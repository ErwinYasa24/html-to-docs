# HTML to DOCX Service

Python FastAPI microservice that converts uploaded HTML files into DOCX documents using Pandoc via `pypandoc`.

## Features
- Upload `.html`/`.htm` files and receive a generated `.docx` download
- Lightweight API with `/health` and `/convert` endpoints
- Built-in payload validation and background cleanup of temporary files
- Math-aware conversion. Tag `<span class="math-tex">` otomatis diubah ke delimiters TeX (`\\( … \\)`), lalu diekspor sebagai persamaan Word
- Streamlit UI sederhana: upload → convert → download dalam satu halaman

## Requirements
- Python 3.11+
- Pandoc installed and available on `PATH`
- (Optional) MathJax CDN access when rendering math content inside Word

## Setup
```bash
cd html_to_docx_service
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Verify Pandoc:
```bash
pandoc --version
```

## Running the FastAPI Service
```bash
# opsi 1
uvicorn app.main:app --reload

# opsi 2
python -m app.main

# opsi 3
python run_api.py
```

Upload files by sending `multipart/form-data` request ke `POST /convert`:
```bash
curl -X POST \
  -F "html_file=@sample.html" \
  http://localhost:8000/convert \
  -o output.docx
```

Docs are available at `http://localhost:8000/docs` when running in dev mode.

## Running the Streamlit UI
```bash
streamlit run streamlit_app.py
```

Antarmuka web akan menyediakan form unggah HTML dan tombol unduh DOCX hasil konversi.

## Project Layout
```
html_to_docx_service/
├── app/
│   ├── __init__.py
│   ├── converter.py   # Pandoc-powered conversion helpers
│   └── main.py        # FastAPI app definition
├── README.md
└── requirements.txt

Tambahan:
- `streamlit_app.py` – Antarmuka Streamlit untuk konversi HTML → DOCX.
```

## Testing Ideas
- Upload HTML that includes images, tables, and math spans to confirm formatting is preserved
- Simulate empty uploads or wrong MIME types to verify error handling
- Run load tests if expecting high concurrency (consider a shared temp directory cleaner)
