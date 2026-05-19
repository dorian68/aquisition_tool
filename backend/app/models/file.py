import uuid
from datetime import datetime

from sqlalchemy import Column, DateTime, ForeignKey, Integer, String
from sqlalchemy.orm import relationship

from app.models.database import Base


def new_uuid() -> str:
    return str(uuid.uuid4())


class GeneratedFile(Base):
    __tablename__ = "generated_files"

    id = Column(String, primary_key=True, default=new_uuid)
    dashboard_id = Column(String, ForeignKey("dashboards.id"), nullable=False, index=True)
    user_id = Column(String, ForeignKey("users.id"), nullable=False, index=True)
    file_type = Column(String, nullable=False, default="xlsx")
    storage_path = Column(String, nullable=False)
    download_count = Column(Integer, nullable=False, default=0)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    expires_at = Column(DateTime, nullable=True)

    dashboard = relationship("Dashboard", back_populates="generated_files")

