import uuid
from datetime import datetime

from sqlalchemy import Column, DateTime, ForeignKey, JSON, String
from sqlalchemy.orm import relationship

from app.models.database import Base


def new_uuid() -> str:
    return str(uuid.uuid4())


class Dashboard(Base):
    __tablename__ = "dashboards"

    id = Column(String, primary_key=True, default=new_uuid)
    upload_id = Column(String, ForeignKey("uploads.id"), nullable=False, index=True)
    user_id = Column(String, ForeignKey("users.id"), nullable=True, index=True)
    dashboard_type = Column(String, nullable=False)
    spec_json = Column(JSON, nullable=False)
    preview_json = Column(JSON, nullable=False)
    status = Column(String, nullable=False, default="ready")
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    generated_files = relationship("GeneratedFile", back_populates="dashboard")

