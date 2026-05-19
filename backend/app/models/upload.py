import uuid
from datetime import datetime

from sqlalchemy import Column, DateTime, Float, ForeignKey, Integer, JSON, String
from sqlalchemy.orm import relationship

from app.models.database import Base


def new_uuid() -> str:
    return str(uuid.uuid4())


class Upload(Base):
    __tablename__ = "uploads"

    id = Column(String, primary_key=True, default=new_uuid)
    user_id = Column(String, ForeignKey("users.id"), nullable=True, index=True)
    original_filename = Column(String, nullable=False)
    source_name = Column(String, nullable=True)
    user_session_id = Column(String, nullable=True, index=True)
    storage_path = Column(String, nullable=False)
    file_size = Column(Integer, nullable=False)
    delimiter = Column(String, nullable=False)
    row_count = Column(Integer, nullable=False)
    column_count = Column(Integer, nullable=False)
    status = Column(String, nullable=False, default="uploaded")
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    profiles = relationship("DatasetProfile", back_populates="upload")


class DatasetProfile(Base):
    __tablename__ = "dataset_profiles"

    id = Column(String, primary_key=True, default=new_uuid)
    upload_id = Column(String, ForeignKey("uploads.id"), nullable=False, index=True)
    profile_json = Column(JSON, nullable=False)
    quality_score = Column(Float, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

    upload = relationship("Upload", back_populates="profiles")

