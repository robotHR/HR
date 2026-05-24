from datetime import datetime
from sqlalchemy import Column, Integer, String, Text, DateTime, Boolean

from app.core.database import Base


class JobPost(Base):
    __tablename__ = "job_posts"

    id = Column(Integer, primary_key=True, index=True)
    titlu = Column(String, nullable=False)
    departament = Column(String, nullable=True)
    locatie = Column(String, nullable=True)
    tip_contract = Column(String, nullable=True, default="Full-time")
    descriere = Column(Text, nullable=True)
    responsabilitati = Column(Text, nullable=True)
    cerinte = Column(Text, nullable=True)
    ce_oferim = Column(Text, nullable=True)
    activ = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
