"""Utilities for converting HTML content into DOCX documents."""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET
from typing import Iterable, Sequence

import pypandoc

from app.preprocess import prepare_html


class InvalidHtmlError(ValueError):
    """Raised when the supplied HTML payload is empty or malformed."""


class PandocNotInstalledError(RuntimeError):
    """Raised when Pandoc is unavailable on the host system."""


class ConversionFailedError(RuntimeError):
    """Raised when Pandoc fails to convert the provided HTML."""


@dataclass
class ConversionResult:
    """Holds metadata for a finished conversion job."""

    output_path: Path
    download_name: str
    workdir: Path


class HtmlToDocxConverter:
    """Perform HTML to DOCX conversions using Pandoc."""

    _WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    _STYLE_IDS = ("Heading1", "Heading2", "Heading3", "Title")
    _DEFAULT_FONT_COLOR = "000000"

    def __init__(
        self,
        *,
        pandoc_args: Sequence[str] | None = None,
        input_format: str | None = None,
        auto_install_pandoc: bool = True,
    ) -> None:
        # Enable TeX math detection inside HTML (e.g. \( ... \) or $$ ... $$)
        self._input_format = input_format or "html+tex_math_dollars+tex_math_single_backslash"
        self._pandoc_args: Sequence[str] = pandoc_args or ("--mathjax",)

        if auto_install_pandoc:
            self._ensure_pandoc_available()

    def convert_input_bytes(self, payload: bytes, original_name: str | None = None) -> ConversionResult:
        """Convert HTML or DOCX payload into a DOCX file.

        Args:
            payload: HTML document as bytes.
            original_name: Optional original filename for naming the output.

        Returns:
            ConversionResult with file paths and display name.

        Raises:
            InvalidHtmlError: If the payload is empty after trimming whitespace.
            PandocNotInstalledError: If Pandoc is not available.
            ConversionFailedError: If Pandoc fails to convert the payload.
        """

        if not payload or not payload.strip():
            raise InvalidHtmlError("Uploaded content is empty.")

        workdir = Path(tempfile.mkdtemp(prefix="html2docx_"))
        output_stem = self._sanitize_filename(original_name)

        input_extension = self._detect_extension(original_name)
        input_path = workdir / f"input{input_extension}"

        if input_extension in {".html", ".htm"}:
            promote = self._should_promote_entities(payload)
            processed_payload = prepare_html(payload, promote_entities=promote)
            input_path.write_bytes(processed_payload)
            src_format = self._input_format
        elif input_extension == ".docx":
            input_path.write_bytes(payload)
            # Convert docx -> html to normalize math, then back to docx
            html_path = workdir / "intermediate.html"
            pypandoc.convert_file(
                str(input_path),
                "html",
                outputfile=str(html_path),
                extra_args=list(self._pandoc_args),
            )
            normalized_html = prepare_html(html_path.read_bytes(), promote_entities=True)
            html_path.write_bytes(normalized_html)
            input_path = html_path
            src_format = self._input_format
        else:
            raise InvalidHtmlError("Unsupported file type. Upload HTML or DOCX.")

        output_name = f"{output_stem or 'document'}.docx"
        output_path = workdir / output_name

        try:
            pypandoc.convert_file(
                str(input_path),
                "docx",
                format=src_format,
                outputfile=str(output_path),
                extra_args=list(self._pandoc_args),
            )
        except OSError as exc:  # Raised when Pandoc binary is missing
            raise PandocNotInstalledError(
                "Pandoc is required for conversion. Install Pandoc and ensure it is on PATH."
            ) from exc
        except RuntimeError as exc:
            raise ConversionFailedError(f"Pandoc failed to convert HTML: {exc}") from exc

        if not output_path.exists():
            raise ConversionFailedError("Pandoc reported success but no DOCX file was created.")

        self._apply_style_overrides(output_path)

        return ConversionResult(output_path=output_path, download_name=output_name, workdir=workdir)

    def convert_to_flat_html(self, payload: bytes, original_name: str | None = None) -> str:
        """Convert payload to DOCX then back to sanitized HTML without document boilerplate."""

        result = self.convert_input_bytes(payload, original_name=original_name)

        try:
            html_output = pypandoc.convert_file(
                str(result.output_path),
                "html",
                extra_args=list(self._pandoc_args),
            )
        except OSError as exc:
            raise PandocNotInstalledError(
                "Pandoc is required for conversion. Install Pandoc and ensure it is on PATH."
            ) from exc
        except RuntimeError as exc:
            raise ConversionFailedError(f"Pandoc failed to convert DOCX to HTML: {exc}") from exc
        finally:
            self.cleanup([result.workdir])

        cleaned = prepare_html(html_output.encode("utf-8"), promote_entities=True)
        return cleaned.decode("utf-8")

    @staticmethod
    def cleanup(paths: Iterable[Path | str]) -> None:
        """Remove temporary files or directories created during conversion."""

        for target in paths:
            try:
                shutil.rmtree(target)  # Handles directories
            except NotADirectoryError:
                Path(target).unlink(missing_ok=True)
            except FileNotFoundError:
                continue

    @staticmethod
    def _sanitize_filename(name: str | None) -> str:
        """Return a filesystem-safe filename."""

        if not name:
            return "document.html"
        sanitized = re.sub(r"[^A-Za-z0-9_.-]+", "_", name)
        sanitized = sanitized.strip("._") or "document"
        return sanitized

    @staticmethod
    def _ensure_html_extension(name: str) -> str:
        return name if name.lower().endswith((".html", ".htm")) else f"{name}.html"

    @staticmethod
    def _detect_extension(name: str | None) -> str:
        if not name:
            return ".html"
        lowered = name.lower()
        for ext in (".html", ".htm", ".docx"):
            if lowered.endswith(ext):
                return ext
        return ".html"

    @staticmethod
    def _ensure_pandoc_available() -> None:
        """Ensure pandoc binary is available, downloading if necessary."""

        try:
            pypandoc.get_pandoc_path()
        except OSError:
            try:
                pypandoc.download_pandoc()
            except OSError as exc:
                raise PandocNotInstalledError(
                    "Pandoc is required for conversion and automatic download failed."
                ) from exc

    def _apply_style_overrides(self, docx_path: Path) -> None:
        """Enforce consistent heading colours on the generated DOCX."""

        try:
            self._force_heading_colors(docx_path, self._DEFAULT_FONT_COLOR)
        except Exception:
            # Styling failures shouldn't break conversion; keep the document as-is.
            return

    def _force_heading_colors(self, docx_path: Path, color_hex: str) -> None:
        with tempfile.TemporaryDirectory(prefix="html2docx_styles_") as tmpdir:
            tmp_root = Path(tmpdir)
            with zipfile.ZipFile(docx_path) as docx_zip:
                docx_zip.extractall(tmp_root)

            styles_path = tmp_root / "word" / "styles.xml"
            if not styles_path.exists():
                return

            ET.register_namespace("w", self._WORD_NS)
            tree = ET.parse(styles_path)
            root = tree.getroot()

            ns = {"w": self._WORD_NS}

            for style_id in self._STYLE_IDS:
                style = root.find(f".//w:style[@w:styleId='{style_id}']", ns)
                if style is None:
                    continue
                self._update_style_color(style, color_hex, ns)

            tree.write(styles_path, encoding="utf-8", xml_declaration=True)

            tmp_docx = tmp_root / "__styled.docx"
            with zipfile.ZipFile(tmp_docx, "w", compression=zipfile.ZIP_DEFLATED) as new_docx:
                for file in tmp_root.rglob("*"):
                    if file == tmp_docx or file.is_dir():
                        continue
                    arcname = file.relative_to(tmp_root).as_posix()
                    new_docx.write(file, arcname)

            shutil.move(tmp_docx, docx_path)

    def _update_style_color(self, style_element: ET.Element, color_hex: str, ns: dict[str, str]) -> None:
        """Ensure both paragraph and run properties force the requested colour."""

        color_tag = f"{{{self._WORD_NS}}}color"

        def _ensure_color(parent: ET.Element) -> None:
            if parent is None:
                return
            color = parent.find("w:color", ns)
            if color is None:
                color = ET.SubElement(parent, color_tag)
            color.set(f"{{{self._WORD_NS}}}val", color_hex)
            for attr in ("themeColor", "themeTint", "themeShade"):
                color.attrib.pop(f"{{{self._WORD_NS}}}{attr}", None)

        paragraph_props = style_element.find("w:pPr", ns)
        if paragraph_props is None:
            paragraph_props = ET.SubElement(style_element, f"{{{self._WORD_NS}}}pPr")
        run_props = paragraph_props.find("w:rPr", ns)
        if run_props is None:
            run_props = ET.SubElement(paragraph_props, f"{{{self._WORD_NS}}}rPr")
        _ensure_color(run_props)

        style_run_props = style_element.find("w:rPr", ns)
        if style_run_props is None:
            style_run_props = ET.SubElement(style_element, f"{{{self._WORD_NS}}}rPr")
        _ensure_color(style_run_props)

    @staticmethod
    def _should_promote_entities(payload: bytes) -> bool:
        snippet = payload[:2048].lower()
        return b"&lt;" in snippet and b"&gt;" in snippet
