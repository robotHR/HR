from datetime import datetime
from sqlalchemy import Column, Integer, String, DateTime, Boolean

from app.core.database import Base


class ScorecardToken(Base):
    __tablename__ = "scorecard_tokens"

    id = Column(Integer, primary_key=True, index=True)

    token = Column(String, unique=True, nullable=False, index=True)
    candidate_id = Column(Integer, nullable=False, index=True)
    interview_id = Column(Integer, nullable=True)

    evaluator_email = Column(String, nullable=True)
    evaluator_name = Column(String, nullable=True)

    used = Column(Boolean, default=False)
    expires_at = Column(DateTime, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
