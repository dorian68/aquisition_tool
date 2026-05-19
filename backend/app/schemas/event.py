from datetime import datetime
from typing import Any

from pydantic import BaseModel


class EventResponse(BaseModel):
    id: str
    event_type: str
    user_id: str | None = None
    payload_json: dict[str, Any]
    created_at: datetime


class EventsListResponse(BaseModel):
    events: list[EventResponse]

