from sqlalchemy import Column,Integer,String,Text
from app.core.database import Base

class Candidate(Base):
    __tablename__ = "candidates"

    id = Column(Integer, primary_key=True)
    name = Column(String)
    email = Column(String)
    phone = Column(String)
    city = Column(String)
    skills = Column(Text)
    experience = Column(Text)
    score = Column(Integer, default=0)
    status = Column(String, default="Nou")
    summary = Column(Text)
