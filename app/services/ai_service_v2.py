"""
AI SERVICE V2 - 7 MULTI-CRITERIA SCORING
STRICT FACTS ONLY - No hallucinations
Uses extracted CV data + job profile to calculate scores
"""

import json
import os
from typing import Dict, List, Optional
from cv_extractor_v2 import CVExtractor
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()


class MultiCriteriaScorer:
    """Score candidates against job profiles using 7 strict criteria"""
    
    def __init__(self, job_profile: Dict):
        self.job_profile = job_profile
        self.client = OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=os.getenv("OPENROUTER_API_KEY")
        )
    
    def clean_score(self, value) -> int:
        """Ensure score is 0-100"""
        try:
            return max(0, min(100, int(value)))
        except:
            return 0
    
    def score_direct_experience_match(self, cv_data: Dict) -> int:
        """
        SCORE 1: Direct Experience Match
        Does the CV explicitly mention the target job position?
        
        95-100: Exact match ("casier")
        70-94: Related but slightly different ("cashier" vs "casier")
        40-69: Related field (e.g., "customer service" for casier)
        1-39: Tangentially related
        0: No mention at all
        """
        target_job = self.job_profile.get("job_title", "").lower()
        must_have = self.job_profile.get("must_have", [])
        
        positions = cv_data.get("positions", [])
        skills = cv_data.get("skills", [])
        cv_text = cv_data.get("cv_text_preview", "").lower()
        
        # Check if exact position mentioned
        for pos in positions:
            pos_text = str(pos.get("position", "")).lower()
            if target_job in pos_text or pos_text in target_job:
                return 95  # Exact match
        
        # Check if must_have skills are in CV
        must_have_found = sum(1 for item in must_have if any(
            skill.lower() in item.lower() or item.lower() in skill.lower()
            for skill in skills
        ))
        
        if must_have_found >= len(must_have) * 0.8:  # 80% of must-haves
            return 75
        elif must_have_found >= len(must_have) * 0.5:  # 50% of must-haves
            return 50
        else:
            return 20  # Very few or no must-haves
    
    def score_years_in_role(self, cv_data: Dict) -> int:
        """
        SCORE 2: Years in Role (Seniority)
        
        95-100: 5+ years
        80-94: 3-5 years
        60-79: 1-3 years
        30-59: <1 year (junior/no experience)
        0: No data / Unknown
        """
        years = cv_data.get("years_in_role")
        
        if years is None:
            return 0  # STRICT: No data = 0, not assumed
        
        if years >= 5:
            return 95
        elif years >= 3:
            return 80
        elif years >= 1:
            return 60
        else:
            return 30
    
    def score_company_level(self, cv_data: Dict) -> int:
        """
        SCORE 3: Company Level
        Where did they work? Major retailer? Startup? Unknown?
        
        90-100: Major international (Carrefour, METRO, Auchan, etc)
        75-89: Large national company
        50-74: Medium company
        25-49: Small company / Self-employed
        0: Unknown (not mentioned)
        """
        last_company = cv_data.get("last_company", "").lower()
        
        major_companies = [
            "carrefour", "metro", "auchan", "lidl", "penny", "kaufland",
            "dhl", "ups", "fedex", "gls", "sameday",
            "accenture", "deloitte", "pwc", "kpmg", "siemens", "bosch"
        ]
        
        large_companies = [
            "altex", "edenred", "brd", "bcr", "bnp", "unicredit",
            "ford", "dacia", "renault", "volvo", "mercedes",
            "emerson", "philips", "asus", "dell"
        ]
        
        if not last_company:
            return 0  # STRICT: Unknown = 0
        
        for major in major_companies:
            if major in last_company:
                return 90
        
        for large in large_companies:
            if large in last_company:
                return 75
        
        # If company mentioned but not in our lists
        if last_company and len(last_company) > 2:
            return 50  # Assume medium-sized
        
        return 0
    
    def score_skill_keywords_present(self, cv_data: Dict) -> int:
        """
        SCORE 4: Skill Keywords Present
        Does CV explicitly mention required technical skills?
        
        Calculate % of must_have skills found in CV
        90-100: 90%+ of must_haves
        75-89: 70-89% of must_haves
        50-74: 40-69% of must_haves
        25-49: 10-39% of must_haves
        0: 0% of must_haves
        """
        must_have = self.job_profile.get("must_have", [])
        skills = cv_data.get("skills", [])
        
        if not must_have:
            return 50  # No skills to check
        
        found_count = 0
        for must in must_have:
            must_lower = must.lower()
            for skill in skills:
                if skill.lower() in must_lower or must_lower in skill.lower():
                    found_count += 1
                    break
        
        percentage = (found_count / len(must_have)) * 100 if must_have else 0
        
        if percentage >= 90:
            return 95
        elif percentage >= 70:
            return 80
        elif percentage >= 40:
            return 60
        elif percentage >= 10:
            return 35
        else:
            return 0
    
    def score_job_stability(self, cv_data: Dict) -> int:
        """
        SCORE 5: Job Stability
        Is there job hopping? Gaps? Long tenure?
        
        STRICT: We can only check if data is available
        90-100: 3+ years at same company (from extracted years_in_role)
        75-89: 1-3 years stable
        50-74: Multiple positions, some stability
        25-49: Job hopping visible or gaps
        0: Unknown
        """
        years = cv_data.get("years_in_role")
        positions = cv_data.get("positions", [])
        
        # If we know tenure length
        if years:
            if years >= 3:
                return 90
            elif years >= 1:
                return 75
            else:
                return 40
        
        # If no years but have position data
        if positions and len(positions) > 0:
            if len(positions) == 1:
                return 70  # Single role = stable
            elif len(positions) <= 2:
                return 55  # Few roles = somewhat stable
            else:
                return 30  # Many roles = job hopping risk
        
        return 0  # Unknown
    
    def score_growth_progression(self, cv_data: Dict) -> int:
        """
        SCORE 6: Growth Progression
        Is there career progression visible? (junior→middle→senior)?
        
        STRICT: Only count explicit progression mentions
        90-100: Clear progression (promoted, advanced role)
        70-89: Some growth (more responsibility visible)
        40-69: Lateral moves (similar roles)
        1-39: No clear progression, but stable
        0: Unknown / Not enough data
        """
        positions = cv_data.get("positions", [])
        cv_text = cv_data.get("cv_text_preview", "").lower()
        
        # Look for progression keywords
        progression_keywords = ["promoted", "promovat", "senior", "lead", "manager", "director", "coordonator", "responsabil"]
        
        has_progression = any(keyword in cv_text for keyword in progression_keywords)
        
        if has_progression and len(positions) >= 2:
            return 80
        elif len(positions) >= 2:
            return 50  # Multiple positions but no clear progression
        elif len(positions) == 1:
            return 30  # Single role = no progression visible
        else:
            return 0  # Unknown
    
    def score_red_flags(self, cv_data: Dict) -> int:
        """
        SCORE 7: Red Flags / CV Quality
        Are there warnings? Gaps, vague descriptions, inconsistencies?
        
        90-100: Clean CV, detailed, no issues
        70-89: Minor issues (small gaps, some vague text)
        40-69: Moderate issues (longer gaps, generic descriptions)
        1-39: Serious issues (major gaps, very vague)
        0: Critical issues (inconsistencies, suspicious patterns)
        """
        cv_length = cv_data.get("cv_length", 0)
        positions = cv_data.get("positions", [])
        email = cv_data.get("email")
        phone = cv_data.get("phone")
        skills = cv_data.get("skills", [])
        
        issues = 0
        max_issues = 5
        
        # Issue: CV too short (probably incomplete)
        if cv_length < 500:
            issues += 2
        
        # Issue: No positions listed
        if not positions or len(positions) == 0:
            issues += 1
        
        # Issue: No contact info
        if not email and not phone:
            issues += 1
        
        # Issue: No skills listed
        if not skills or len(skills) == 0:
            issues += 1
        
        # Issue: Missing years of experience
        if cv_data.get("years_in_role") is None:
            issues += 0.5
        
        # Calculate score inversely
        score = 90 - (issues * 15)
        return self.clean_score(score)
    
    def calculate_final_score(self, subscores: Dict) -> int:
        """
        Calculate FINAL SCORE as weighted average of 7 subscores
        All subscores weighted equally (14.28% each)
        """
        weights = {
            "direct_experience": 0.20,       # 20% - most important
            "years_in_role": 0.15,           # 15%
            "company_level": 0.12,           # 12%
            "skill_keywords": 0.15,          # 15%
            "job_stability": 0.12,           # 12%
            "growth_progression": 0.08,      # 8%
            "red_flags": 0.08,               # 8% (reversed - lower is worse)
        }
        
        final = sum(
            subscores.get(key, 0) * weight
            for key, weight in weights.items()
        )
        
        return self.clean_score(final)
    
    def score_candidate_for_job(self, cv_data: Dict) -> Dict:
        """
        MAIN FUNCTION: Score candidate against job profile
        Returns all 7 subscores + final score
        """
        subscores = {
            "direct_experience": self.score_direct_experience_match(cv_data),
            "years_in_role": self.score_years_in_role(cv_data),
            "company_level": self.score_company_level(cv_data),
            "skill_keywords": self.score_skill_keywords_present(cv_data),
            "job_stability": self.score_job_stability(cv_data),
            "growth_progression": self.score_growth_progression(cv_data),
            "red_flags": self.score_red_flags(cv_data),
        }
        
        final_score = self.calculate_final_score(subscores)
        
        # Determine recommendation
        if final_score >= 80:
            recommendation = "TOP CANDIDATE"
        elif final_score >= 65:
            recommendation = "POTENTIAL CANDIDATE"
        elif final_score >= 50:
            recommendation = "CONSIDER"
        else:
            recommendation = "LOW FIT"
        
        return {
            "candidate_name": cv_data.get("name", "Unknown"),
            "email": cv_data.get("email"),
            "phone": cv_data.get("phone"),
            "final_score": final_score,
            "recommendation": recommendation,
            "subscores": subscores,
            "last_company": cv_data.get("last_company"),
            "years_in_role": cv_data.get("years_in_role"),
            "skills": cv_data.get("skills"),
            "driving_license": cv_data.get("driving_license"),
        }
