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

st.set_page_config(page_title="HTML → DOCX Converter", page_icon="📝")
st.markdown(
    "<h1 style='text-align: center;'>HTML → DOCX Converter</h1>",
    unsafe_allow_html=True,
)
st.write(
    """<p style='text-align:center;'>Unggah berkas `.html` atau `.docx`, lalu unduh hasil konversi `.docx` dengan rumus yang sudah diperbaiki.</p>""",
    unsafe_allow_html=True,
)

converter = HtmlToDocxConverter()

with st.form("upload-form"):
    uploaded_file = st.file_uploader(
        "Pilih berkas HTML atau DOCX",
        type=["html", "htm", "docx"],
        accept_multiple_files=False,
    )
    col_left, col_center, col_right = st.columns([3, 2, 3])
    with col_center:
        submit_btn = st.form_submit_button("Convert", use_container_width=True)

if submit_btn:
    if not uploaded_file:
        st.warning("Silakan pilih berkas HTML atau DOCX terlebih dahulu.")
    else:
        try:
            raw_content = uploaded_file.read()
            result = converter.convert_input_bytes(raw_content, original_name=uploaded_file.name)
        except InvalidHtmlError as exc:
            st.error(f"Berkas tidak valid: {exc}")
        except PandocNotInstalledError as exc:
            st.error(f"Pandoc belum terpasang: {exc}")
        except ConversionFailedError as exc:
            st.error(f"Konversi gagal: {exc}")
        else:
            docx_bytes = result.output_path.read_bytes()
            st.success("Konversi berhasil. Klik tombol di bawah untuk mengunduh DOCX baru.")
            st.download_button(
                label="Download DOCX",
                data=io.BytesIO(docx_bytes),
                file_name=result.download_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            HtmlToDocxConverter.cleanup([result.workdir])
