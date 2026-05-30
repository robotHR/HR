from sqlalchemy import Column, Integer, String, DateTime, Boolean
from datetime import datetime, timedelta

from app.core.database import Base


class SchedulingToken(Base):
    __tablename__ = "scheduling_tokens"

    id = Column(Integer, primary_key=True, index=True)
    token = Column(String, unique=True, index=True, nullable=False)
    candidate_id = Column(Integer, nullable=False, index=True)
    expires_at = Column(DateTime, nullable=False)
    used = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
