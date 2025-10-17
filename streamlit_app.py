"""Streamlit UI for uploading HTML files and downloading converted DOCX output."""
from __future__ import annotations

import io

import streamlit as st

from app.converter import (
    ConversionFailedError,
    HtmlToDocxConverter,
    InvalidHtmlError,
    PandocNotInstalledError,
)
from app.preprocess import normalize_math_spans

st.set_page_config(page_title="HTML → DOCX Converter", page_icon="📝")
st.markdown(
    "<h1 style='text-align: center;'>HTML → DOCX Converter</h1>",
    unsafe_allow_html=True,
)
st.write(
    """<p style='text-align:center;'>Unggah berkas `.html` lalu unduh hasil konversi `.docx`.</p>""",
    unsafe_allow_html=True,
)

converter = HtmlToDocxConverter()

with st.form("upload-form"):
    uploaded_file = st.file_uploader(
        "Pilih berkas HTML",
        type=["html", "htm"],
        accept_multiple_files=False,
    )
    col_left, col_center, col_right = st.columns([3, 2, 3])
    with col_center:
        submit_btn = st.form_submit_button("Convert", use_container_width=True)

if submit_btn:
    if not uploaded_file:
        st.warning("Silakan pilih berkas HTML terlebih dahulu.")
    else:
        try:
            raw_content = uploaded_file.read()
            normalized = normalize_math_spans(raw_content.decode("utf-8", errors="ignore")).encode("utf-8")
            result = converter.convert_html_bytes(normalized, original_name=uploaded_file.name)
        except InvalidHtmlError as exc:
            st.error(f"Berkas tidak valid: {exc}")
        except PandocNotInstalledError as exc:
            st.error(f"Pandoc belum terpasang: {exc}")
        except ConversionFailedError as exc:
            st.error(f"Konversi gagal: {exc}")
        else:
            docx_bytes = result.output_path.read_bytes()
            st.success("Konversi berhasil. Klik tombol di bawah untuk mengunduh.")
            st.download_button(
                label="Download DOCX",
                data=io.BytesIO(docx_bytes),
                file_name=result.download_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            HtmlToDocxConverter.cleanup([result.workdir])
