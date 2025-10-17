"""Helpers to normalize HTML prior to conversion."""
from __future__ import annotations

import re

MATH_SPAN_PATTERN = re.compile(
    r"<span[^>]*class=\"[^\"]*math-tex[^\"]*\"[^>]*>(.*?)</span>",
    flags=re.IGNORECASE | re.DOTALL,
)


def normalize_math_spans(html: str) -> str:
    """Replace math span wrappers with Pandoc-friendly TeX delimiters."""

    def _replace(match: re.Match[str]) -> str:
        inner = match.group(1).strip()
        if inner.startswith("$") or inner.startswith("\\("):
            return inner
        return f"\\({inner}\\)"

    return MATH_SPAN_PATTERN.sub(_replace, html)


def prepare_html(payload: bytes) -> bytes:
    """Apply preprocessing transforms to raw HTML bytes."""

    text = payload.decode("utf-8", errors="ignore")
    text = normalize_math_spans(text)
    return text.encode("utf-8")
