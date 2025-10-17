"""Utilities for converting HTML content into DOCX documents."""
from __future__ import annotations

import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
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
            processed_payload = prepare_html(payload)
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
            normalized_html = prepare_html(html_path.read_bytes())
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

        return ConversionResult(output_path=output_path, download_name=output_name, workdir=workdir)

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
