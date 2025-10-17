"""Helpers to normalize HTML prior to conversion."""
from __future__ import annotations

import html as html_module
import re

DOCTYPE_PATTERN = re.compile(r"<!DOCTYPE[^>]*>", re.IGNORECASE)
ENCODED_DOCTYPE_PATTERN = re.compile(r"&lt;!DOCTYPE[^&]*&gt;", re.IGNORECASE)
STYLE_BLOCK_PATTERN = re.compile(r"<style[^>]*>.*?</style>", re.IGNORECASE | re.DOTALL)
ENCODED_STYLE_PATTERN = re.compile(r"&lt;style[^&]*&gt;.*?&lt;/style&gt;", re.IGNORECASE | re.DOTALL)
MATH_SPAN_PATTERN = re.compile(
    r"<span[^>]*class=\"[^\"]*math-tex[^\"]*\"[^>]*>(.*?)</span>",
    flags=re.IGNORECASE | re.DOTALL,
)

TEXT_NODE_PATTERN = re.compile(r">([^<>]+)<")

LATEX_KEYWORD_PATTERN = re.compile(
    r"\\(frac|times|sqrt|sum|prod|int|left|right|binom|over|cdot|dots|ldots|sin|cos|tan|log|ln|pi|alpha|beta|gamma|theta)",
    re.IGNORECASE,
)

PROMOTABLE_TAGS = {
    "a",
    "abbr",
    "article",
    "aside",
    "b",
    "blockquote",
    "body",
    "br",
    "button",
    "caption",
    "code",
    "col",
    "colgroup",
    "dd",
    "details",
    "div",
    "dl",
    "dt",
    "em",
    "figcaption",
    "figure",
    "footer",
    "form",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "head",
    "header",
    "hr",
    "html",
    "link",
    "meta",
    "img",
    "input",
    "label",
    "legend",
    "li",
    "main",
    "math",
    "mfrac",
    "mi",
    "mn",
    "mo",
    "mrow",
    "ms",
    "msqrt",
    "msub",
    "msubsup",
    "msup",
    "mtext",
    "mspace",
    "nav",
    "ol",
    "option",
    "p",
    "pre",
    "title",
    "section",
    "select",
    "small",
    "span",
    "strong",
    "sub",
    "summary",
    "sup",
    "table",
    "tbody",
    "td",
    "textarea",
    "tfoot",
    "th",
    "thead",
    "tr",
    "u",
    "ul",
}

ESCAPED_TAG_PATTERN = re.compile(
    r"&lt;(?P<full>(?P<prefix>/?)\s*(?P<tag>[A-Za-z][A-Za-z0-9:-]*)(?P<tail>[^<>]*?))&gt;",
    flags=re.IGNORECASE,
)

EMBEDDED_HTML_PATTERN = re.compile(r"&lt;html.*?&lt;/html&gt;", re.IGNORECASE | re.DOTALL)

MSO_SPAN_PATTERN = re.compile(r"</?span[^>]*>", re.IGNORECASE)
MSO_CLASS_PATTERN = re.compile(r'\sclass="?(?:Mso[^"\s>]*)"?', re.IGNORECASE)
BLOCK_IN_P_PATTERN = re.compile(
    r"<p>\s*(<(?:html|head|body|div|section|article|aside|nav|main|figure|figcaption|h[1-6]|ul|ol|li|table|thead|tbody|tfoot|tr|td|th|blockquote|pre)[^>]*>.*?</(?:html|head|body|div|section|article|aside|nav|main|figure|figcaption|h[1-6]|ul|ol|li|table|thead|tbody|tfoot|tr|td|th|blockquote|pre)>)\s*</p>",
    re.IGNORECASE | re.DOTALL,
)


def normalize_math_spans(html: str) -> str:
    """Replace math span wrappers with Pandoc-friendly TeX delimiters."""

    def _replace(match: re.Match[str]) -> str:
        inner = match.group(1).strip()
        if inner.startswith("$") or inner.startswith("\\("):
            return inner
        return f"\\({inner}\\)"

    return MATH_SPAN_PATTERN.sub(_replace, html)


def promote_escaped_html(html: str) -> str:
    """Unescape HTML-like entities that should become real markup."""

    if "&lt;" not in html or "&gt;" not in html:
        return html

    def _promote(match: re.Match[str]) -> str:
        tag = match.group("tag")
        if tag.lower() not in PROMOTABLE_TAGS:
            return match.group(0)

        body_unescaped = html_module.unescape(match.group(0))
        cleaned = body_unescaped.strip()

        if not cleaned.startswith("<"):
            return match.group(0)

        if cleaned.endswith(">"):
            return cleaned
        return f"{cleaned}>"

    return ESCAPED_TAG_PATTERN.sub(_promote, html)


def strip_html_boilerplate(html: str) -> str:
    """Remove DOCTYPE declarations and style blocks from HTML."""

    html = DOCTYPE_PATTERN.sub("", html)
    html = ENCODED_DOCTYPE_PATTERN.sub("", html)
    html = STYLE_BLOCK_PATTERN.sub("", html)
    html = ENCODED_STYLE_PATTERN.sub("", html)
    return html


def extract_embedded_html(html: str) -> tuple[str, bool]:
    match = EMBEDDED_HTML_PATTERN.search(html)
    if not match:
        return html, False
    return html_module.unescape(match.group(0)), True


def simplify_word_export(html: str) -> str:
    """Collapse Word-export wrappers so embedded HTML becomes valid markup."""

    if "Mso" not in html and "mso-" not in html:
        return html

    html = MSO_SPAN_PATTERN.sub("", html)
    html = MSO_CLASS_PATTERN.sub("", html)

    def _strip_wrappers(value: str) -> str:
        while True:
            new_value = BLOCK_IN_P_PATTERN.sub(r"\1", value)
            if new_value == value:
                break
            value = new_value
        return value

    html = _strip_wrappers(html)
    html = re.sub(
        r"(</?(?:html|head|body|div|section|article|aside|nav|main|figure|figcaption|h[1-6]|ul|ol|li|table|thead|tbody|tfoot|tr|td|th|blockquote|pre)[^>]*>)\s*</p>",
        r"\1",
        html,
        flags=re.IGNORECASE,
    )
    html = re.sub(
        r"(</?(?:meta|title|link)[^>]*>)\s*</p>",
        r"\1",
        html,
        flags=re.IGNORECASE,
    )
    html = re.sub(r"<p>\s*</p>", "", html)
    html = re.sub(
        r"<p>\s*(<(?:meta|title|div|section|article|aside|nav|main|figure|figcaption|h[1-6]|ul|ol|li|table|thead|tbody|tfoot|tr|td|th|blockquote|pre|body|html)[^>]*>)",
        r"\1",
        html,
        flags=re.IGNORECASE,
    )
    html = re.sub(r"(</(?:div|section|article|aside|nav|main|figure|figcaption|ul|ol|table|tbody|tfoot|thead|tr|td|th|blockquote|pre|body|html)>)\s*</p>", r"\1", html, flags=re.IGNORECASE)
    html = html.replace("<p></div>", "</div>")
    html = html.replace("<p></head>", "</head>")
    html = html.replace("<p></body>", "</body>")
    html = html.replace("<p></html>", "</html>")
    html = html.replace("<p><p", "<p").replace("</p></p>", "</p>")
    return html


def prepare_html(payload: bytes, *, promote_entities: bool = False) -> bytes:
    """Apply preprocessing transforms to raw HTML bytes."""

    text = payload.decode("utf-8", errors="ignore")
    text, embedded_found = extract_embedded_html(text)
    if embedded_found:
        text = simplify_word_export(text)
    text = strip_html_boilerplate(text)
    if promote_entities or embedded_found:
        text = promote_escaped_html(text)
        text = strip_html_boilerplate(text)
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
