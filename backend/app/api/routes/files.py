from datetime import datetime

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, status
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from app.core.security import get_current_user
from app.models.database import get_db
from app.models.file import GeneratedFile
from app.models.user import User
from app.services.event_logger import log_event
from app.services.n8n_client import N8nClient
from app.services.storage import LocalStorage

router = APIRouter(prefix="/files", tags=["files"])


@router.get("/{file_id}/download")
def download_file(
    file_id: str,
    background_tasks: BackgroundTasks,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> FileResponse:
    generated_file = db.get(GeneratedFile, file_id)
    if not generated_file:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="File not found")
    if generated_file.user_id != current_user.id:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="File does not belong to current user")
    if generated_file.expires_at and generated_file.expires_at < datetime.utcnow():
        raise HTTPException(status_code=status.HTTP_410_GONE, detail="File link expired")

    path = LocalStorage().path_for(generated_file.storage_path)
    if not path.exists():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="File missing from storage")

    generated_file.download_count += 1
    db.add(generated_file)
    payload = {
        "event": "file_downloaded",
        "user": {"id": current_user.id, "email": current_user.email},
        "file": {"file_id": generated_file.id, "dashboard_id": generated_file.dashboard_id},
        "timestamp": datetime.utcnow().isoformat(),
    }
    db.commit()
    log_event(db, "file_downloaded", payload, current_user.id)
    background_tasks.add_task(N8nClient().send_event, "file_downloaded", payload)

    return FileResponse(
        path=path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"optiquant-dashboard-{generated_file.id}.xlsx",
    )

