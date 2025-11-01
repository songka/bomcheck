from __future__ import annotations

from functools import lru_cache

MODES = ("t2s", "s2t")

try:  # pragma: no cover - optional dependency
    from opencc import OpenCC
except ImportError:  # pragma: no cover
    OpenCC = None  # type: ignore[misc, assignment]


@lru_cache(maxsize=None)
def _get_converter(mode: str) -> OpenCC | None:
    if OpenCC is None:
        return None
    try:
        return OpenCC(mode)
    except Exception:  # pragma: no cover - opencc failure
        return None


def _prepare_value(value: str | object) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def normalize_text(value: str) -> str:
    base = _prepare_value(value)
    converter = _get_converter("t2s")
    if converter:
        try:
            converted = converter.convert(base)
            if converted:
                return converted.strip().lower()
        except Exception:  # pragma: no cover - opencc failure
            pass
    return base


def normalized_variants(value: str) -> set[str]:
    variants: set[str] = set()
    base = _prepare_value(value)
    if base:
        variants.add(base)
    for mode in MODES:
        converter = _get_converter(mode)
        if not converter:
            continue
        try:
            converted = converter.convert(base)
        except Exception:  # pragma: no cover - opencc failure
            continue
        normalized = _prepare_value(converted)
        if normalized:
            variants.add(normalized)
    return variants
