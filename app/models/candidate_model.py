from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, DateTime

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
    job_id = Column(Integer, nullable=True)
    status = Column(String, nullable=True)
    cv_file = Column(String, nullable=True)

    batch_id = Column(String, nullable=True)
    visible_in_dashboard = Column(Integer, default=1)

    # ── Campuri agent AI decizie ──────────────────────────────────────────────
    decision_reason          = Column(Text,   nullable=True)   # De ce a primit scorul asta
    missing_requirements     = Column(Text,   nullable=True)   # Ce lipseste concret din CV
    must_verify_by_phone     = Column(Text,   nullable=True)   # Ce verifici la telefon
    recommended_next_action  = Column(String, nullable=True)   # CALL_NOW|CALL_LATER|KEEP_FOR_OTHER_ROLE|REJECT
    priority                 = Column(String, nullable=True)   # HIGH|MEDIUM|LOW
    manager_summary          = Column(Text,   nullable=True)   # Rezumat pentru manager (3 randuri)
    phone_call_script        = Column(Text,   nullable=True)   # Script gata de folosit la telefon
    interview_questions      = Column(Text,   nullable=True)   # Intrebari recomandate pentru interviu
    better_role_match        = Column(String, nullable=True)   # Alt rol mai potrivit pentru candidat
    reject_reason_internal   = Column(Text,   nullable=True)   # Motiv respingere (intern, nu se trimite)

    created_at = Column(DateTime, default=datetime.utcnow)
