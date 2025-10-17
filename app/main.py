"""FastAPI application for HTML to DOCX conversion."""
from __future__ import annotations

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.converter import (
    ConversionFailedError,
    HtmlToDocxConverter,
    InvalidHtmlError,
    PandocNotInstalledError,
)
from app.preprocess import prepare_html

ALLOWED_CONTENT_TYPES = {"text/html", "application/xhtml+xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024  # 5 MB default cap

app = FastAPI(title="HTML to DOCX Converter", version="0.1.0")
converter = HtmlToDocxConverter()


@app.get("/health")
async def health_check() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/convert")
async def convert_html_document(
    background_tasks: BackgroundTasks,
    html_file: UploadFile = File(..., description="HTML file to convert to DOCX"),
) -> FileResponse:
    if html_file.content_type not in ALLOWED_CONTENT_TYPES and not html_file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=415, detail="Unsupported file type. Upload HTML or DOCX.")

    payload = await html_file.read()
    if len(payload) > MAX_FILE_SIZE_BYTES:
        raise HTTPException(status_code=413, detail="File too large. Max 5 MB allowed.")

    try:
        is_html = html_file.filename.lower().endswith((".html", ".htm"))
        processed = prepare_html(payload) if is_html else payload
        result = converter.convert_input_bytes(processed, original_name=html_file.filename)
    except InvalidHtmlError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except PandocNotInstalledError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except ConversionFailedError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    background_tasks.add_task(HtmlToDocxConverter.cleanup, [result.workdir])

    return FileResponse(
        path=str(result.output_path),
        filename=result.download_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=True)
