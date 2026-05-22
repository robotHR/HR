from sqlalchemy import Column, Integer, String, Text

from app.core.database import Base


class Candidate(Base):
    __tablename__ = "candidates"

    id = Column(Integer, primary_key=True, index=True)

    name = Column(String, nullable=True)
    email = Column(String, nullable=True)
    phone = Column(String, nullable=True)
    position = Column(String, nullable=True)
    experience = Column(String, nullable=True)

    skills = Column(Text, nullable=True)
    companies = Column(Text, nullable=True)

    score = Column(Integer, nullable=True)
    level = Column(String, nullable=True)

    strengths = Column(Text, nullable=True)
    weaknesses = Column(Text, nullable=True)
    summary = Column(Text, nullable=True)

    job_title = Column(String, nullable=True)
    status = Column(String, nullable=True)
    cv_file = Column(String, nullable=True)

    batch_id = Column(String, nullable=True)
    visible_in_dashboard = Column(Integer, default=1)
