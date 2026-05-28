from sqlalchemy import Column, Integer, String, DateTime, Text
from datetime import datetime

from app.core.database import Base


class Interview(Base):
    __tablename__ = "interviews"

    id = Column(Integer, primary_key=True, index=True)

    candidate_id = Column(Integer, nullable=False, index=True)
    candidate_name = Column(String, nullable=True)   # denormalizat pentru viteza
    candidate_email = Column(String, nullable=True)
    job_title = Column(String, nullable=True)

    data = Column(String, nullable=False)            # "2026-06-10"
    ora = Column(String, nullable=False)             # "10:00"
    locatie = Column(String, nullable=True)
    durata = Column(Integer, default=60)             # minute

    status = Column(String, default="programat")     # programat / finalizat / anulat

    created_at = Column(DateTime, default=datetime.utcnow)
