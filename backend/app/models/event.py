import uuid
from datetime import datetime

from sqlalchemy import Column, DateTime, ForeignKey, Integer, JSON, String, Text

from app.models.database import Base


def new_uuid() -> str:
    return str(uuid.uuid4())


class Event(Base):
    __tablename__ = "events"

    id = Column(String, primary_key=True, default=new_uuid)
    event_type = Column(String, index=True, nullable=False)
    user_id = Column(String, ForeignKey("users.id"), nullable=True, index=True)
    payload_json = Column(JSON, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class N8nWebhookLog(Base):
    __tablename__ = "n8n_webhook_logs"

    id = Column(String, primary_key=True, default=new_uuid)
    event_type = Column(String, index=True, nullable=False)
    payload_json = Column(JSON, nullable=False)
    response_status = Column(Integer, nullable=True)
    response_body = Column(Text, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

