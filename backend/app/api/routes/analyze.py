from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from app.models.database import get_db
from app.models.upload import DatasetProfile, Upload
from app.schemas.upload import DatasetProfileResponse
from app.services.csv_loader import CSVLoader, CSVValidationError
from app.services.dataset_profiler import DatasetProfiler
from app.services.storage import LocalStorage

router = APIRouter(prefix="/uploads", tags=["analysis"])


@router.post("/{upload_id}/analyze", response_model=DatasetProfileResponse)
def analyze_upload(upload_id: str, db: Session = Depends(get_db)) -> DatasetProfileResponse:
    upload = db.get(Upload, upload_id)
    if not upload or upload.status == "deleted":
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Upload not found")

    try:
        content = LocalStorage().read(upload.storage_path)
        loaded = CSVLoader().load_bytes(upload.original_filename, content, "text/csv")
    except (CSVValidationError, ValueError) as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc

    dataset_profile = DatasetProfiler().profile(loaded.dataframe)
    profile = DatasetProfile(
        upload_id=upload.id,
        profile_json=dataset_profile,
        quality_score=dataset_profile["quality_score"],
    )
    db.add(profile)
    upload.status = "analyzed"
    db.add(upload)
    db.commit()

    return DatasetProfileResponse(dataset_profile=dataset_profile)

