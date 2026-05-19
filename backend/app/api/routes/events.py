from fastapi import APIRouter, Depends, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from app.models.database import get_db
from app.models.event import Event
from app.schemas.event import EventsListResponse

router = APIRouter(prefix="/events", tags=["events"])


@router.get("", response_model=EventsListResponse)
def list_events(limit: int = Query(default=50, ge=1, le=200), db: Session = Depends(get_db)) -> EventsListResponse:
    events = db.execute(select(Event).order_by(Event.created_at.desc()).limit(limit)).scalars().all()
    return EventsListResponse(
        events=[
            {
                "id": event.id,
                "event_type": event.event_type,
                "user_id": event.user_id,
                "payload_json": event.payload_json,
                "created_at": event.created_at,
            }
            for event in events
        ]
    )

