import uuid
from datetime import datetime, timedelta

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.core.config import get_settings
from app.core.security import get_current_user
from app.models.dashboard import Dashboard
from app.models.database import get_db
from app.models.file import GeneratedFile
from app.models.upload import DatasetProfile, Upload
from app.models.user import Lead, User
from app.schemas.dashboard import DashboardSpecResponse, GenerateXlsxResponse, PreviewResponse
from app.services.csv_loader import CSVLoader
from app.services.dashboard_planner import DashboardPlanner
from app.services.dashboard_preview import DashboardPreviewBuilder
from app.services.dataset_profiler import DatasetProfiler
from app.services.event_logger import log_event
from app.services.n8n_client import N8nClient
from app.services.storage import LocalStorage
from app.services.xlsx_generator import XlsxDashboardGenerator

router = APIRouter(tags=["dashboards"])


@router.post("/uploads/{upload_id}/dashboard-spec", response_model=DashboardSpecResponse)
def create_dashboard_spec(upload_id: str, db: Session = Depends(get_db)) -> DashboardSpecResponse:
    upload = db.get(Upload, upload_id)
    if not upload or upload.status == "deleted":
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Upload not found")

    df = _load_upload_dataframe(upload)
    dataset_profile = _latest_profile(db, upload.id)
    if dataset_profile is None:
        dataset_profile = DatasetProfiler().profile(df)
        db.add(DatasetProfile(upload_id=upload.id, profile_json=dataset_profile, quality_score=dataset_profile["quality_score"]))
        db.commit()

    spec = DashboardPlanner().plan(dataset_profile, df)
    preview_data = DashboardPreviewBuilder().build(df, spec, dataset_profile)
    dashboard = Dashboard(
        id=spec["dashboard_id"],
        upload_id=upload.id,
        dashboard_type=spec["dashboard_type"],
        spec_json=spec,
        preview_json=preview_data,
        status="ready",
    )
    db.add(dashboard)
    upload.status = "dashboard_ready"
    db.add(upload)
    db.commit()

    return DashboardSpecResponse(**spec)


@router.get("/dashboards/{dashboard_id}/preview", response_model=PreviewResponse)
def dashboard_preview(dashboard_id: str, db: Session = Depends(get_db)) -> PreviewResponse:
    dashboard = db.get(Dashboard, dashboard_id)
    if not dashboard:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Dashboard not found")
    return PreviewResponse(dashboard_spec=dashboard.spec_json, preview_data=dashboard.preview_json)


@router.post("/dashboards/{dashboard_id}/generate-xlsx", response_model=GenerateXlsxResponse)
def generate_xlsx(
    dashboard_id: str,
    background_tasks: BackgroundTasks,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> GenerateXlsxResponse:
    dashboard = db.get(Dashboard, dashboard_id)
    if not dashboard:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Dashboard not found")

    upload = db.get(Upload, dashboard.upload_id)
    if not upload or upload.status == "deleted":
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Upload not found")

    df = _load_upload_dataframe(upload)
    dataset_profile = _latest_profile(db, upload.id) or DatasetProfiler().profile(df)
    xlsx_bytes = XlsxDashboardGenerator().generate(df, dashboard.spec_json, dashboard.preview_json, dataset_profile)

    file_id = str(uuid.uuid4())
    storage_path = LocalStorage().save_generated_file(file_id, xlsx_bytes)
    settings = get_settings()
    generated_file = GeneratedFile(
        id=file_id,
        dashboard_id=dashboard.id,
        user_id=current_user.id,
        file_type="xlsx",
        storage_path=storage_path,
        expires_at=datetime.utcnow() + timedelta(days=settings.generated_file_ttl_days),
    )
    dashboard.user_id = current_user.id
    db.add(generated_file)
    db.add(dashboard)
    _ensure_lead(db, current_user, upload.id, dashboard.id)
    db.commit()

    payload = {
        "event": "dashboard_generated",
        "user": {"id": current_user.id, "email": current_user.email, "name": current_user.name},
        "dashboard": {"dashboard_id": dashboard.id, "dashboard_type": dashboard.dashboard_type},
        "file": {"file_id": generated_file.id, "download_url": f"/api/v1/files/{generated_file.id}/download"},
    }
    log_event(db, "dashboard_generated", payload, current_user.id)
    background_tasks.add_task(N8nClient().send_event, "dashboard_generated", payload)

    return GenerateXlsxResponse(file_id=generated_file.id, download_url=f"/api/v1/files/{generated_file.id}/download")


def _load_upload_dataframe(upload: Upload):
    content = LocalStorage().read(upload.storage_path)
    return CSVLoader().load_bytes(upload.original_filename, content, "text/csv").dataframe


def _latest_profile(db: Session, upload_id: str) -> dict | None:
    profile = db.execute(
        select(DatasetProfile).where(DatasetProfile.upload_id == upload_id).order_by(DatasetProfile.created_at.desc())
    ).scalars().first()
    return profile.profile_json if profile else None


def _ensure_lead(db: Session, user: User, upload_id: str, dashboard_id: str) -> None:
    existing = db.execute(select(Lead).where(Lead.user_id == user.id, Lead.dashboard_id == dashboard_id)).scalars().first()
    if existing:
        return
    db.add(
        Lead(
            user_id=user.id,
            email=user.email,
            name=user.name,
            source="csv_dashboard_generator",
            first_upload_id=upload_id,
            dashboard_id=dashboard_id,
        )
    )

