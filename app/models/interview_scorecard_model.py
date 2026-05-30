from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, DateTime

from app.core.database import Base


class InterviewScorecard(Base):
    __tablename__ = "interview_scorecards"

    id = Column(Integer, primary_key=True, index=True)

    candidate_id = Column(Integer, nullable=False, index=True)
    interview_id = Column(Integer, nullable=True, index=True)

    evaluator_name = Column(String, nullable=True)

    technical_skills_score     = Column(Integer, nullable=True)
    attention_to_detail_score  = Column(Integer, nullable=True)
    communication_score        = Column(Integer, nullable=True)
    team_fit_score             = Column(Integer, nullable=True)

    overall_recommendation = Column(String, nullable=True)  # Strong Hire / Hire / No Hire
    comments = Column(Text, nullable=True)

    created_at = Column(DateTime, default=datetime.utcnow)
