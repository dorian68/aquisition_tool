from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile, status
from sqlalchemy.orm import Session

from app.models.database import get_db
from app.models.upload import Upload
from app.schemas.upload import DeleteUploadResponse, UploadCsvResponse
from app.services.csv_loader import CSVLoader, CSVValidationError
from app.services.storage import LocalStorage

router = APIRouter(prefix="/uploads", tags=["uploads"])


@router.post("/csv", response_model=UploadCsvResponse, status_code=status.HTTP_201_CREATED)
async def upload_csv(
    file: UploadFile = File(...),
    user_session_id: str | None = Form(default=None),
    source_name: str | None = Form(default=None),
    db: Session = Depends(get_db),
) -> UploadCsvResponse:
    content = await file.read()
    loader = CSVLoader()
    try:
        loaded = loader.load_bytes(file.filename or "upload.csv", content, file.content_type)
    except CSVValidationError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc

    storage_path = LocalStorage().save_upload(file.filename or "upload.csv", content)
    upload = Upload(
        original_filename=file.filename or "upload.csv",
        source_name=source_name,
        user_session_id=user_session_id,
        storage_path=storage_path,
        file_size=len(content),
        delimiter=loaded.delimiter,
        row_count=loaded.rows,
        column_count=loaded.columns,
        status="uploaded",
    )
    db.add(upload)
    db.commit()
    db.refresh(upload)

    return UploadCsvResponse(
        upload_id=upload.id,
        status=upload.status,
        rows=upload.row_count,
        columns=upload.column_count,
        detected_delimiter=upload.delimiter,
    )


@router.delete("/{upload_id}", response_model=DeleteUploadResponse)
def delete_upload(upload_id: str, db: Session = Depends(get_db)) -> DeleteUploadResponse:
    upload = db.get(Upload, upload_id)
    if not upload:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Upload not found")

    try:
        LocalStorage().delete(upload.storage_path)
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc

    upload.status = "deleted"
    db.add(upload)
    db.commit()
    return DeleteUploadResponse(upload_id=upload.id, status=upload.status)

