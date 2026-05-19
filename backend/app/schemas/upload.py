from typing import Any

from pydantic import BaseModel, Field


class UploadCsvResponse(BaseModel):
    upload_id: str
    status: str
    rows: int
    columns: int
    detected_delimiter: str
    message: str = "CSV uploaded successfully"


class ColumnProfile(BaseModel):
    name: str
    type: str
    semantic_type: str | None = None
    missing_rate: float
    unique_count: int
    sample_values: list[Any] = Field(default_factory=list)


class DatasetProfileResponse(BaseModel):
    dataset_profile: dict[str, Any]


class DeleteUploadResponse(BaseModel):
    upload_id: str
    status: str

