from typing import Any

from sqlalchemy.orm import Session

from app.models.event import Event


def log_event(db: Session, event_type: str, payload: dict[str, Any], user_id: str | None = None) -> Event:
    event = Event(event_type=event_type, user_id=user_id, payload_json=payload)
    db.add(event)
    db.commit()
    db.refresh(event)
    return event

