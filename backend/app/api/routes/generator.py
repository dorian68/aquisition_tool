from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path
from typing import Literal

from fastapi import APIRouter, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

from app.core.config import get_settings
from app.services.premium_dashboard_generator import (
    PremiumDashboardGenerationError,
    generate_premium_dashboard,
)

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
    vba_project_file: UploadFile | None = File(default=None, description="Optional compiled vbaProject.bin for XLSM output."),
) -> FileResponse:
    content = await file.read()
    max_bytes = get_settings().max_upload_mb * 1024 * 1024
    if not content:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Uploaded CSV is empty")
    if len(content) > max_bytes:
        raise HTTPException(status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE, detail="Uploaded CSV exceeds max size")

    temp_dir = Path(tempfile.mkdtemp(prefix="optiquant_dashboard_"))
    try:
        csv_path = temp_dir / "input.csv"
        csv_path.write_bytes(content)

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
        )
    except PremiumDashboardGenerationError as exc:
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
