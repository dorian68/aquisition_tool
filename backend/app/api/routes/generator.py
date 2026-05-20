from __future__ import annotations

import json
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Literal

from fastapi import APIRouter, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

from app.core.config import get_settings
from app.services.ai_analyst import prepare_python_analysis, serialized_context_length, validate_csv_shape
from app.services.premium_dashboard_generator import (
    PremiumDashboardGenerationError,
    generate_premium_dashboard,
)
from app.services.storage import LocalStorage

router = APIRouter(prefix="/generator", tags=["generator"])

TemplateSlug = Literal["auto", "dark-saas", "fintech-executive", "light-consulting"]
OutputFormat = Literal["xlsx", "xlsm"]

MEDIA_TYPES = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
}


@router.post("/dashboard")
async def generate_dashboard_file(
    file: UploadFile = File(..., description="CSV file to transform into an Excel dashboard."),
    template: TemplateSlug = Form(default="auto"),
    output_format: OutputFormat = Form(default="xlsx"),
    client_name: str | None = Form(default=None),
    hide_settings: bool = Form(default=False),
    include_ai_analysis: bool = Form(default=True),
    analysis_id: str | None = Form(default=None),
    vba_project_file: UploadFile | None = File(default=None, description="Optional compiled vbaProject.bin for XLSM output."),
) -> FileResponse:
    settings = get_settings()
    content = await _read_valid_csv_upload(file)

    temp_dir = Path(tempfile.mkdtemp(prefix="optiquant_dashboard_"))
    try:
        csv_path = temp_dir / "input.csv"
        csv_path.write_bytes(content)

        ai_report: dict | None = None
        if include_ai_analysis:
            if analysis_id:
                validate_csv_shape(csv_path, settings=settings)
                ai_report = _load_stored_ai_report(analysis_id)
            else:
                analysis = prepare_python_analysis(csv_path, template=template, output_format=output_format, settings=settings)
                ai_report = analysis["ai_report"]
        else:
            validate_csv_shape(csv_path, settings=settings)

        vba_project_path: Path | None = None
        if output_format == "xlsm" and vba_project_file is not None:
            vba_content = await vba_project_file.read()
            if not vba_content:
                raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Uploaded vbaProject.bin is empty")
            vba_project_path = temp_dir / "vbaProject.bin"
            vba_project_path.write_bytes(vba_content)

        output_path = temp_dir / f"dashboard.{output_format}"
        metadata = generate_premium_dashboard(
            csv_path,
            output_path,
            template=template,
            vba_project_path=vba_project_path,
            client_name=client_name,
            hide_settings=hide_settings,
            ai_report=ai_report,
        )
    except PremiumDashboardGenerationError as exc:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    except ValueError as exc:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise

    selected = str(metadata["selected_template"])
    filename = f"dashboard-{selected}.{output_format}"
    response_headers = {
        "X-Dashboard-Template": selected,
        "X-Dashboard-Dataset-Type": str(metadata.get("dataset_type") or ""),
        "X-Dashboard-Metadata": json.dumps(metadata, ensure_ascii=True),
    }

    return FileResponse(
        path=Path(metadata["output"]),
        media_type=MEDIA_TYPES[output_format],
        filename=filename,
        headers=response_headers,
        background=BackgroundTask(shutil.rmtree, temp_dir, ignore_errors=True),
    )


@router.post("/analyze")
async def analyze_dashboard_file(
    file: UploadFile = File(..., description="CSV file to analyze without sending raw rows to the AI model."),
    template: TemplateSlug = Form(default="auto"),
    output_format: OutputFormat = Form(default="xlsx"),
) -> JSONResponse:
    settings = get_settings()
    content = await _read_valid_csv_upload(file)
    temp_dir = Path(tempfile.mkdtemp(prefix="optiquant_analysis_"))
    try:
        csv_path = temp_dir / "input.csv"
        csv_path.write_bytes(content)
        analysis = prepare_python_analysis(csv_path, template=template, output_format=output_format, settings=settings)
        ai_context = analysis["ai_context"]
        ai_report = analysis["ai_report"]
        ai_metadata = analysis["ai_metadata"]
        analysis_id = str(uuid.uuid4())
        payload = {
            "analysis_id": analysis_id,
            "recommended_template": ai_context["dashboard_context"]["recommended_template"],
            "selected_template": analysis["selected_template"],
            "dataset_overview": ai_context["dataset_overview"],
            "data_quality": ai_context["data_quality"],
            "ai_report": ai_report,
            "limits": {
                "raw_rows_sent_to_llm": 0,
                "ai_context_chars": serialized_context_length(ai_context),
                "max_ai_context_chars": settings.max_ai_context_chars,
            },
            "ai_metadata": {
                "provider": "langgraph",
                "model": ai_metadata.get("model"),
                "used_fallback": bool(ai_metadata.get("used_fallback")),
            },
        }
        LocalStorage(settings).save_spec_json(analysis_id, {
            **payload,
            "ai_context": ai_context,
            "template": template,
            "output_format": output_format,
        })
        return JSONResponse(payload)
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


async def _read_valid_csv_upload(file: UploadFile) -> bytes:
    content = await file.read()
    max_bytes = get_settings().max_upload_mb * 1024 * 1024
    if not content:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Uploaded CSV is empty")
    if len(content) > max_bytes:
        raise HTTPException(status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE, detail="Uploaded CSV exceeds max size")
    suffix = Path(file.filename or "upload.csv").suffix.lower()
    if suffix and suffix != ".csv":
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Only CSV uploads are supported")
    return content


def _load_stored_ai_report(analysis_id: str) -> dict:
    try:
        payload = LocalStorage().read_spec_json(analysis_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Analysis not found") from exc
    ai_report = payload.get("ai_report")
    if not isinstance(ai_report, dict):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Stored analysis is invalid")
    return ai_report
