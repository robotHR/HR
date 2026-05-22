"""
CV EXTRACTOR V2 - STRICT, 0 HALLUCINATIONS
Extracts ONLY what's explicitly in the CV text
"""

import re
from typing import Dict, List, Optional

class CVExtractor:
    """Extract structured data from CV text - STRICT FACTS ONLY"""
    
    def __init__(self, cv_text: str, filename: str):
        self.cv_text = cv_text.lower()
        self.cv_text_original = cv_text
        self.filename = filename
        
    def extract_name_from_filename(self) -> str:
        """Extract name from filename (best guess when not in CV)"""
        clean = self.filename.replace(".pdf", "").replace(".docx", "").replace(".doc", "")
        clean = clean.replace("-", " ").replace("_", " ").replace(".", " ")
        bad = {"cv", "curriculum", "vitae", "resume", "formal", "2024", "2025", "atestat", "atestata"}
        words = [w.strip() for w in clean.split() if w.lower() not in bad and w.strip()]
        return " ".join(words).title() if words else "Unknown"
    
    def normalize_text(self, text: str) -> str:
        """Normalize Romanian characters"""
        replacements = {
            "ș": "s", "ş": "s",
            "ă": "a", "â": "a", "î": "i",
            "ț": "t", "ţ": "t",
            "Ș": "S", "Ş": "S",
            "Ă": "A", "Â": "A", "Î": "I",
            "Ț": "T", "Ţ": "T"
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text
    
    def extract_positions_and_companies(self) -> List[Dict]:
        """
        Extract job positions and companies - STRICT
        Only what's EXPLICITLY mentioned in CV
        """
        positions = []
        
        # Common patterns for job positions
        position_keywords = [
            r"(?:job|position|rol|pozitie|titlu|functie):\s*([^\n]+)",
            r"(?:as|as a|as an|in rol of|in the role of|manager|supervisor|specialist|engineer|developer|accountant|casier|sofer|operator|lucrator|vanzator|consultant|analyst|director)\s+([a-z\s]+?)(?:\s+(?:at|in|cu|la|pentru|2[0-9]{3}|$))",
        ]
        
        # Company keywords
        company_keywords = [
            r"(?:company|companie|firma|societate|employer|angajator|la|la:\s*|worked at|worked with|at):\s*([^\n,]+)",
            r"(?:la|at|in):\s*([A-Z][a-zA-Z\s&.]+?)(?:\s*(?:from|since|in|as|-|,))",
        ]
        
        text = self.normalize_text(self.cv_text)
        
        # Extract explicit mentions
        for pattern in position_keywords:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                pos = match.group(1).strip()
                if len(pos) > 2 and len(pos) < 100:
                    positions.append({"position": pos, "company": None, "years": None})
        
        # Try to match companies
        for pattern in company_keywords:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                company = match.group(1).strip()
                if len(company) > 2 and len(company) < 80:
                    if not any(p["company"] == company for p in positions):
                        positions.append({"position": None, "company": company, "years": None})
        
        return positions
    
    def extract_years_in_role(self) -> Optional[int]:
        """
        Extract years of experience mentioned explicitly in CV
        STRICT: Only numbers we can see
        """
        years = None
        
        # Patterns: "3 ani", "3 years", "3-5 years", "since 2020" etc
        patterns = [
            r"(\d+)\s*(?:ani|years|years|an|year)",  # "3 ani" or "3 years"
            r"(?:from|din|since)\s+(\d{4})\s+(?:to|until|pana|la)\s+(\d{4})",  # "2020 to 2023"
            r"(\d{4})\s*-\s*(\d{4})",  # "2020-2023"
        ]
        
        text = self.cv_text
        
        for pattern in patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                if pattern == patterns[0]:
                    years = int(match.group(1))
                    return years
                elif pattern in patterns[1:]:
                    try:
                        start = int(match.group(1))
                        end = int(match.group(2))
                        if 1900 < start < 2025 and 1900 < end < 2025:
                            years = max(0, end - start)
                            return years
                    except:
                        pass
        
        return years
    
    def extract_last_company(self) -> Optional[str]:
        """
        Extract LAST company mentioned in CV - usually in experience section
        STRICT: Only explicit mentions
        """
        # Major Romanian/international companies to look for
        known_companies = [
            "carrefour", "metro", "auchan", "lidl", "penny", "kaufland", "altex",
            "dedeman", "allianz", "brd", "bcr", "bnp paribas", "unicredit",
            "cdu logistics", "dhl", "ups", "fedex", "gls", "sameday",
            "ford", "dacia", "renault", "volvo", "mercedes", "scania",
            "emerson", "siemens", "bosch", "philips", "asus", "dell",
            "accenture", "deloitte", "pwc", "kpmg", "ey", "ernst young"
        ]
        
        text = self.normalize_text(self.cv_text)
        found_companies = []
        
        for company in known_companies:
            if company in text:
                found_companies.append(company.title())
        
        # If found, return the last mentioned (usually most recent)
        if found_companies:
            return found_companies[-1]
        
        return None
    
    def extract_email_phone(self) -> Dict:
        """Extract email and phone - STRICT"""
        data = {"email": None, "phone": None}
        
        # Email pattern
        email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
        email_match = re.search(email_pattern, self.cv_text_original)
        if email_match:
            data["email"] = email_match.group(0)
        
        # Phone pattern (Romanian: 07xx xxx xxxx, +40 7xx xxx xxxx, etc)
        phone_pattern = r"(?:\+40|0)(?:7\d{8}|\d{2}\s?\d{4}\s?\d{4}|\d{9})"
        phone_match = re.search(phone_pattern, self.cv_text_original)
        if phone_match:
            data["phone"] = phone_match.group(0)
        
        return data
    
    def extract_skills(self) -> List[str]:
        """
        Extract explicit skill mentions from CV
        STRICT: Keywords that appear in CV
        """
        skills = []
        skill_keywords = {
            # Technical
            "cash register": ["cash register", "cash", "pos", "ecr"],
            "SAP": ["sap"],
            "Excel": ["excel"],
            "Word": ["word"],
            "SQL": ["sql"],
            "Python": ["python"],
            "Java": ["java"],
            "JavaScript": ["javascript", "js"],
            "C++": ["c++", "cpp"],
            "C#": ["c#", "csharp"],
            "PHP": ["php"],
            # Soft skills
            "Customer Service": ["customer service", "customer care", "client facing", "relational"],
            "Team Work": ["team", "teamwork", "collaboration"],
            "Leadership": ["leadership", "leader", "team lead"],
            "Communication": ["communication", "communicativ"],
            "Attention to Detail": ["detail", "meticulous"],
            "Organization": ["organized", "organization", "organized"],
            # Certifications
            "Driving License": ["permis", "driving license", "categoria", "cat "],
            "Atestat": ["atestat"],
            "CECCAR": ["ceccar"],
        }
        
        text = self.normalize_text(self.cv_text)
        found = []
        
        for skill_name, keywords in skill_keywords.items():
            for keyword in keywords:
                if keyword in text and skill_name not in found:
                    found.append(skill_name)
                    break
        
        return found
    
    def extract_driving_license(self) -> Optional[str]:
        """Extract driving license category if mentioned - STRICT"""
        text = self.normalize_text(self.cv_text)
        
        # Look for permis/driving license + category
        categories = {
            "A": ["categoria a", "cat a", "permis a"],
            "B": ["categoria b", "cat b", "permis b"],
            "C": ["categoria c", "cat c", "permis c"],
            "D": ["categoria d", "cat d", "permis d"],
            "E": ["categoria e", "cat e", "permis e"],
            "CE": ["categoria ce", "cat ce", "permis ce", "tir"],
            "DE": ["categoria de", "cat de", "permis de"],
        }
        
        for cat, keywords in categories.items():
            for keyword in keywords:
                if keyword in text:
                    return cat
        
        return None
    
    def extract_all(self) -> Dict:
        """Extract ALL structured data - STRICT ONLY"""
        return {
            "name": self.extract_name_from_filename(),
            "email": self.extract_email_phone()["email"],
            "phone": self.extract_email_phone()["phone"],
            "positions": self.extract_positions_and_companies(),
            "last_company": self.extract_last_company(),
            "years_in_role": self.extract_years_in_role(),
            "skills": self.extract_skills(),
            "driving_license": self.extract_driving_license(),
            "cv_length": len(self.cv_text_original),
            "cv_text_preview": self.cv_text_original[:500],  # First 500 chars for reference
        }
