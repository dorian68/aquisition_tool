from datetime import datetime

from pydantic import BaseModel


class GeneratedFileResponse(BaseModel):
    id: str
    dashboard_id: str
    file_type: str
    download_count: int
    created_at: datetime
    expires_at: datetime | None = None

