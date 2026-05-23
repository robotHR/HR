from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, DateTime, UniqueConstraint

from app.core.database import Base


class CvAnalysis(Base):
    """
    Cache pentru analizele AI.
    Fiecare combinatie (cv_file, job_title_normalized) se analizeaza o singura data.
    A doua oara cand se cauta acelasi job cu acelasi CV -> rezultat instant din DB.
    """
    __tablename__ = "cv_analyses"

    id = Column(Integer, primary_key=True, index=True)
    cv_file = Column(String, nullable=False, index=True)
    job_title_normalized = Column(String, nullable=False, index=True)
    result_json = Column(Text, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    __table_args__ = (
        UniqueConstraint("cv_file", "job_title_normalized", name="uq_cv_job"),
    )
