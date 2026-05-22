"""
MULTI-JOB CV ANALYZER V2
Processes multiple jobs + multiple CVs
Outputs rankings per job + cross-job recommendations
"""

import os
import json
import csv
from typing import Dict, List, Optional
from pathlib import Path
from cv_extractor_v2 import CVExtractor
from ai_service_v2 import MultiCriteriaScorer
import pdfplumber
from docx import Document

# Constants
UPLOAD_FOLDER = "app/uploads"
JOB_PROFILES_PATH = "config/job_profiles.json"


class CVProcessor:
    """Process CV file and extract text"""
    
    @staticmethod
    def extract_pdf_text(pdf_path: str) -> str:
        """Extract text from PDF"""
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        except Exception as e:
            print(f"Error reading PDF {pdf_path}: {e}")
        return text
    
    @staticmethod
    def extract_docx_text(docx_path: str) -> str:
        """Extract text from DOCX"""
        text = ""
        try:
            doc = Document(docx_path)
            for paragraph in doc.paragraphs:
                if paragraph.text:
                    text += paragraph.text + "\n"
        except Exception as e:
            print(f"Error reading DOCX {docx_path}: {e}")
        return text
    
    @staticmethod
    def extract_text_from_file(file_path: str) -> str:
        """Extract text from CV file (PDF, DOCX, DOC)"""
        if file_path.lower().endswith(".pdf"):
            return CVProcessor.extract_pdf_text(file_path)
        elif file_path.lower().endswith(".docx"):
            return CVProcessor.extract_docx_text(file_path)
        return ""


class JobProfileMatcher:
    """Match target job to job profile"""
    
    def __init__(self, job_profiles_path: str):
        self.profiles = self._load_profiles(job_profiles_path)
    
    def _load_profiles(self, path: str) -> List[Dict]:
        """Load job profiles from JSON"""
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("profiles", [])
        except Exception as e:
            print(f"Error loading job profiles: {e}")
            return []
    
    def _normalize_text(self, text: str) -> str:
        """Normalize for matching"""
        text = text.lower()
        replacements = {
            "ș": "s", "ş": "s", "ă": "a", "â": "a",
            "î": "i", "ț": "t", "ţ": "t"
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return " ".join(text.split())
    
    def find_profile(self, target_job: str) -> Optional[Dict]:
        """Find matching job profile"""
        target_normalized = self._normalize_text(target_job)
        
        for profile in self.profiles:
            # Check main title
            if self._normalize_text(profile.get("job_title", "")) == target_normalized:
                return profile
            
            # Check aliases
            for alias in profile.get("aliases", []):
                if self._normalize_text(alias) == target_normalized:
                    return profile
            
            # Partial match
            if target_normalized in self._normalize_text(profile.get("job_title", "")):
                return profile
        
        # If no match found, create fallback profile
        return self._create_fallback_profile(target_job)
    
    def _create_fallback_profile(self, job_title: str) -> Dict:
        """Create basic profile for unknown jobs"""
        return {
            "job_id": "unknown",
            "job_title": job_title,
            "domain": "Unknown Domain",
            "level": "Middle",
            "must_have": ["experience in " + job_title.lower()],
            "nice_to_have": [],
            "reject_if_missing": [],
            "aliases": [],
            "red_flags": [],
        }


class MultiJobAnalyzer:
    """Main analyzer for multiple jobs + CVs"""
    
    def __init__(self, upload_folder: str = UPLOAD_FOLDER, job_profiles_path: str = JOB_PROFILES_PATH):
        self.upload_folder = upload_folder
        self.cv_processor = CVProcessor()
        self.job_matcher = JobProfileMatcher(job_profiles_path)
        self.results = {}
    
    def analyze_all_cvs_for_jobs(self, job_titles: List[str]) -> Dict:
        """
        Main function: Analyze all CVs in folder against multiple jobs
        
        Returns:
        {
            "job_title": [
                {
                    "candidate_name": "...",
                    "score": 85,
                    "recommendation": "TOP CANDIDATE",
                    "subscores": {...}
                }
            ]
        }
        """
        
        results = {job: [] for job in job_titles}
        
        # Get all CV files
        if not os.path.exists(self.upload_folder):
            print(f"Upload folder not found: {self.upload_folder}")
            return results
        
        cv_files = [f for f in os.listdir(self.upload_folder) 
                    if f.lower().endswith((".pdf", ".docx", ".doc"))]
        
        print(f"\n{'='*80}")
        print(f"ANALYZING {len(cv_files)} CVs FOR {len(job_titles)} JOBS")
        print(f"{'='*80}\n")
        
        # Process each CV against each job
        for cv_file in cv_files:
            print(f"\n[CV] {cv_file}")
            cv_path = os.path.join(self.upload_folder, cv_file)
            
            # Extract CV text
            cv_text = self.cv_processor.extract_text_from_file(cv_path)
            if not cv_text:
                print(f"  ❌ Could not extract text from {cv_file}")
                continue
            
            # Extract structured data from CV
            extractor = CVExtractor(cv_text, cv_file)
            cv_data = extractor.extract_all()
            cv_data["cv_text_preview"] = cv_text[:2000]  # Add text preview for scoring
            
            print(f"  ✓ Extracted: {cv_data['name']}")
            if cv_data.get("last_company"):
                print(f"    - Last Company: {cv_data['last_company']}")
            if cv_data.get("years_in_role"):
                print(f"    - Years: {cv_data['years_in_role']}")
            if cv_data.get("skills"):
                print(f"    - Skills: {', '.join(cv_data['skills'][:3])}...")
            
            # Score against each job
            for job_title in job_titles:
                job_profile = self.job_matcher.find_profile(job_title)
                scorer = MultiCriteriaScorer(job_profile)
                score_result = scorer.score_candidate_for_job(cv_data)
                results[job_title].append(score_result)
        
        # Sort each job's candidates by score
        for job_title in job_titles:
            results[job_title].sort(key=lambda x: x.get("final_score", 0), reverse=True)
        
        return results
    
    def print_results(self, results: Dict):
        """Pretty print results"""
        for job_title, candidates in results.items():
            print(f"\n{'='*80}")
            print(f"JOB: {job_title}")
            print(f"{'='*80}")
            
            if not candidates:
                print("No candidates analyzed.")
                continue
            
            for rank, candidate in enumerate(candidates[:10], 1):  # Top 10
                score = candidate.get("final_score", 0)
                name = candidate.get("candidate_name", "Unknown")
                recommendation = candidate.get("recommendation", "N/A")
                company = candidate.get("last_company", "N/A")
                years = candidate.get("years_in_role", "N/A")
                
                # Color coding by score
                if score >= 80:
                    status = "🟢 TOP"
                elif score >= 65:
                    status = "🟡 POTENTIAL"
                else:
                    status = "🔴 LOW FIT"
                
                print(f"\n{rank}. {status} {name}")
                print(f"   Score: {score}/100 | {recommendation}")
                print(f"   Company: {company} | Years: {years}")
                
                # Show subscores
                subscores = candidate.get("subscores", {})
                print(f"   Breakdown:")
                print(f"     - Direct Experience: {subscores.get('direct_experience', 0)}/100")
                print(f"     - Years in Role: {subscores.get('years_in_role', 0)}/100")
                print(f"     - Company Level: {subscores.get('company_level', 0)}/100")
                print(f"     - Skills Match: {subscores.get('skill_keywords', 0)}/100")
                print(f"     - Job Stability: {subscores.get('job_stability', 0)}/100")
                print(f"     - Growth: {subscores.get('growth_progression', 0)}/100")
                print(f"     - CV Quality: {subscores.get('red_flags', 0)}/100")
    
    def export_to_csv(self, results: Dict, output_path: str = "analysis_results.csv"):
        """Export results to CSV with all subscores"""
        rows = []
        
        for job_title, candidates in results.items():
            for rank, candidate in enumerate(candidates, 1):
                row = {
                    "Job": job_title,
                    "Rank": rank,
                    "Candidate": candidate.get("candidate_name", ""),
                    "Email": candidate.get("email", ""),
                    "Phone": candidate.get("phone", ""),
                    "Final_Score": candidate.get("final_score", 0),
                    "Recommendation": candidate.get("recommendation", ""),
                    "Last_Company": candidate.get("last_company", ""),
                    "Years_Experience": candidate.get("years_in_role", ""),
                    "Direct_Experience": candidate.get("subscores", {}).get("direct_experience", 0),
                    "Years_in_Role_Score": candidate.get("subscores", {}).get("years_in_role", 0),
                    "Company_Level": candidate.get("subscores", {}).get("company_level", 0),
                    "Skills_Match": candidate.get("subscores", {}).get("skill_keywords", 0),
                    "Job_Stability": candidate.get("subscores", {}).get("job_stability", 0),
                    "Growth_Progression": candidate.get("subscores", {}).get("growth_progression", 0),
                    "CV_Quality": candidate.get("subscores", {}).get("red_flags", 0),
                }
                rows.append(row)
        
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            if rows:
                writer = csv.DictWriter(f, fieldnames=rows[0].keys())
                writer.writeheader()
                writer.writerows(rows)
        
        print(f"\n✓ Results exported to: {output_path}")
        return output_path
    
    def export_to_json(self, results: Dict, output_path: str = "analysis_results.json"):
        """Export results to JSON"""
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        print(f"✓ Results exported to: {output_path}")
        return output_path


# Example usage
if __name__ == "__main__":
    # Example: Analyze for multiple jobs
    analyzer = MultiJobAnalyzer()
    
    jobs = ["Casier", "Sofer TIR", "Lucrator Depozit"]
    results = analyzer.analyze_all_cvs_for_jobs(jobs)
    
    analyzer.print_results(results)
    analyzer.export_to_csv(results, "analysis_results.csv")
    analyzer.export_to_json(results, "analysis_results.json")
