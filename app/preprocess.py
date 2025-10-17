"""Helpers to normalize HTML prior to conversion."""
from __future__ import annotations

import re

MATH_SPAN_PATTERN = re.compile(
    r"<span[^>]*class=\"[^\"]*math-tex[^\"]*\"[^>]*>(.*?)</span>",
    flags=re.IGNORECASE | re.DOTALL,
)

TEXT_NODE_PATTERN = re.compile(r">([^<>]+)<")

LATEX_KEYWORD_PATTERN = re.compile(
    r"\\(frac|times|sqrt|sum|prod|int|left|right|binom|over|cdot|dots|ldots|sin|cos|tan|log|ln|pi|alpha|beta|gamma|theta)",
    re.IGNORECASE,
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
    text = wrap_bare_latex_sequences(text)
    text = normalize_math_spans(text)
    return text.encode("utf-8")


def wrap_bare_latex_sequences(html: str) -> str:
    """Wrap bare LaTeX-like text nodes in math spans for later normalization."""

    def _wrap_text_node(match: re.Match[str]) -> str:
        original = match.group(1)
        replaced = wrap_bare_latex_in_text(original)
        if replaced == original:
            return match.group(0)
        return f">{replaced}<"

    return TEXT_NODE_PATTERN.sub(_wrap_text_node, html)


def wrap_bare_latex_in_text(text: str) -> str:
    stripped = text.strip()
    if not stripped:
        return text
    if "\\(" in stripped or "\\[" in stripped or "$" in stripped or "<latex>" in stripped:
        return text
    if "\\" not in stripped:
        return text
    if not LATEX_KEYWORD_PATTERN.search(stripped):
        return text

    leading_ws = len(text) - len(text.lstrip())
    trailing_ws = len(text) - len(text.rstrip())

    core = text[leading_ws: len(text) - trailing_ws] if trailing_ws else text[leading_ws:]

    if ":" in core:
        prefix, suffix = core.split(":", 1)
        if "\\" in suffix and not any(token in suffix for token in ("\\(", "\\[", "<span")):
            suffix_ws = len(suffix) - len(suffix.lstrip())
            suffix_core = suffix[suffix_ws:]
            wrapped = f"<span class=\"math-tex\">{suffix_core.strip()}</span>"
            trailing_part = text[len(text) - trailing_ws:] if trailing_ws else ""
            return f"{text[:leading_ws]}{prefix}:{suffix[:suffix_ws]}{wrapped}{trailing_part}"

    wrapped_core = f"<span class=\"math-tex\">{core}</span>"
    trailing_part = text[len(text) - trailing_ws:] if trailing_ws else ""
    return f"{text[:leading_ws]}{wrapped_core}{trailing_part}"
