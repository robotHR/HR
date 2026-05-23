import os
import re
import json
import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pdfplumber

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import win32com.client
except ImportError:
    win32com = None

from dotenv import load_dotenv
from openai import OpenAI
from app.core.database import SessionLocal
from app.models.candidate_model import Candidate
from app.models.cv_analysis_model import CvAnalysis

# Cloudinary este folosit pentru persistenta CV-urilor (vezi cloudinary_service.py)

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

UPLOAD_FOLDER = "app/uploads"
TEMP_FOLDER = "/tmp/nexas_hr"
os.makedirs(TEMP_FOLDER, exist_ok=True)

JOB_PROFILES_PATH = "config/job_profiles.json"
client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=os.getenv("OPENROUTER_API_KEY"))

# ─── CACHE FUNCTIONS ──────────────────────────────────────────────────────────

def _normalize_job_key(job_title):
    """Normalizeaza titlul jobului pentru cache key consistent."""
    return normalize_text(job_title).strip()


def get_cached_analysis(cv_file, job_title):
    """
    Returneaza rezultatul din cache daca exista, altfel None.
    """
    key = _normalize_job_key(job_title)
    db = SessionLocal()
    try:
        row = db.query(CvAnalysis).filter(
            CvAnalysis.cv_file == cv_file,
            CvAnalysis.job_title_normalized == key
        ).first()
        if row:
            return json.loads(row.result_json)
        return None
    except Exception as e:
        logger.warning(f"Cache read error pentru {cv_file}: {e}")
        return None
    finally:
        db.close()


def save_analysis_to_cache(cv_file, job_title, data):
    """
    Salveaza rezultatul analizei in cache.
    Daca exista deja, actualizeaza.
    """
    key = _normalize_job_key(job_title)
    db = SessionLocal()
    try:
        existing = db.query(CvAnalysis).filter(
            CvAnalysis.cv_file == cv_file,
            CvAnalysis.job_title_normalized == key
        ).first()

        result_json = json.dumps(data, ensure_ascii=False)

        if existing:
            existing.result_json = result_json
        else:
            db.add(CvAnalysis(
                cv_file=cv_file,
                job_title_normalized=key,
                result_json=result_json
            ))

        db.commit()
    except Exception as e:
        logger.warning(f"Cache write error pentru {cv_file}: {e}")
        db.rollback()
    finally:
        db.close()


# ─── TEXT EXTRACTION ──────────────────────────────────────────────────────────

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        logger.info(f"Text extras din {pdf_path}: {len(text)} caractere")
    except Exception as e:
        logger.error(f"Eroare la citirea PDF {pdf_path}: {e}")
    return text


def extract_text_from_docx(docx_path):
    text = ""
    if Document is None:
        logger.error("Lipseste python-docx.")
        return text
    try:
        document = Document(docx_path)
        for paragraph in document.paragraphs:
            if paragraph.text:
                text += paragraph.text + "\n"
        for table in document.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    if cell.text:
                        row_text.append(cell.text.strip())
                if row_text:
                    text += " | ".join(row_text) + "\n"
        logger.info(f"Text extras din {docx_path}: {len(text)} caractere")
    except Exception as e:
        logger.error(f"Eroare la citirea DOCX {docx_path}: {e}")
    return text


def extract_text_from_doc(doc_path):
    text = ""
    if win32com is None:
        logger.error("Lipseste pywin32.")
        return text
    word = None
    document = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        document = word.Documents.Open(os.path.abspath(doc_path))
        text = document.Content.Text or ""
        logger.info(f"Text extras din {doc_path}: {len(text)} caractere")
    except Exception as e:
        logger.error(f"Eroare la citirea DOC {doc_path}: {e}")
    finally:
        try:
            if document:
                document.Close(False)
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
    return text


def extract_text_from_file(file_path):
    file_lower = file_path.lower()
    if file_lower.endswith(".pdf"):
        return extract_text_from_pdf(file_path)
    if file_lower.endswith(".docx"):
        return extract_text_from_docx(file_path)
    if file_lower.endswith(".doc"):
        return extract_text_from_doc(file_path)
    return ""


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def normalize_text(value):
    value = str(value or "").lower()
    for a, b in [("ș","s"),("ş","s"),("ă","a"),("â","a"),("î","i"),("ț","t"),("ţ","t"),("."," "),("-"," "),("_"," "),("/"," ")]:
        value = value.replace(a, b)
    return " ".join(value.split())


def load_job_profiles():
    if not os.path.exists(JOB_PROFILES_PATH):
        logger.warning(f"Job profiles file not found at {JOB_PROFILES_PATH}")
        return []
    try:
        with open(JOB_PROFILES_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("profiles", [])
    except Exception as e:
        logger.error(f"Eroare job_profiles.json: {e}")
        return []


def match_job_profile(target_job):
    profiles = load_job_profiles()
    job = normalize_text(target_job)
    best, best_score = None, 0
    for p in profiles:
        candidates = [p.get("job_title",""), p.get("job_id","")] + p.get("aliases", []) + p.get("equivalent_roles", [])
        score = 0
        for item in candidates:
            item = normalize_text(item)
            if not item:
                continue
            if item == job:
                score += 100
            elif item in job or job in item:
                score += 70
            else:
                iw, jw = set(item.split()), set(job.split())
                if iw:
                    score += int((len(iw & jw) / len(iw)) * 45)
        if score > best_score:
            best, best_score = p, score
    return best if best_score >= 25 else None


def build_fallback_profile(target_job):
    return {
        "job_title": target_job, "domain": "General", "level": "Middle",
        "must_have": [], "nice_to_have": [], "reject_if_missing": [],
        "overqualified_risk": ["rector","profesor universitar","director general","CEO","antreprenor"],
        "red_flags": ["lipsa experienta relevanta", "CV generic"],
        "max_score_rules": ["Daca experienta este complet diferita, scor maxim 45."],
        "interview_questions": ["Ce experienta directa ai pentru acest rol?"]
    }


def detect_required_license(target_job, profile=None):
    text = normalize_text(target_job)
    if profile:
        text += " " + normalize_text(" ".join(profile.get("must_have", []) + profile.get("reject_if_missing", [])))
    if "permis ce" in text or "categoria ce" in text or "cat ce" in text or "tir" in text:
        return "CE"
    if "permis b" in text or "categoria b" in text or "cat b" in text or "sofer b" in text:
        return "B"
    if "permis c" in text or "categoria c" in text or "cat c" in text or "sofer c" in text:
        return "C"
    if "permis d" in text or "categoria d" in text or "cat d" in text:
        return "D"
    return ""


def build_prompt(text, target_job):
    profile = match_job_profile(target_job) or build_fallback_profile(target_job)
    profile_json = json.dumps(profile, ensure_ascii=False, indent=2)
    required_license = detect_required_license(target_job, profile)
    license_rule = f"Postul cere permis {required_license}. Nu cere alta categorie." if required_license else "Postul nu cere permis explicit."

    schema = """
{
  "candidate_name": "",
  "email": "",
  "phone": "",
  "companies": [],
  "positions_held": [],
  "current_position": "",
  "recommended_role_for_candidate": "",
  "target_position": "",
  "job_domain": "",
  "years_experience": 0,
  "years_relevant_domain": 0,
  "experience_score": 0,
  "skills_score": 0,
  "soft_skills_score": 0,
  "potential_score": 0,
  "motivation_score": 0,
  "final_score": 0,
  "recommendation": "STRONG_YES|YES|MAYBE|NO|REJECT",
  "confidence_level": "high|medium|low",
  "fit_percentage": 0,
  "overqualification_risk": "low|medium|high",
  "level_mismatch": "low|medium|high",
  "strengths": [],
  "gaps": [],
  "red_flags": [],
  "green_flags": [],
  "transferable_skills": [],
  "growth_potential": "low|medium|high",
  "growth_ceiling": "",
  "interview_questions": [],
  "risk_level": "low|medium|high|critical",
  "retention_probability": 0,
  "wage_alignment": "aligned|higher|lower|unknown",
  "summary": ""
}
"""
    return (
        "TU ESTI RECRUITER HR SENIOR CU 15+ ANI EXPERIENTA IN ROMANIA.\n\n"
        f"ANALIZEZI CV-UL PENTRU POSTUL: {target_job}\n\n"
        "PROFIL JOB DEDICAT DIN config/job_profiles.json:\n" + profile_json + "\n\n"
        f"REGULA PERMIS: {license_rule}\n\n"
        "REGULI OBLIGATORII:\n"
        "- Extrage companiile/institutiile unde a lucrat candidatul si pune-le in companies. Nu inventa companii. Daca nu apar clar, lasa lista goala.\n"
        "- Extrage functiile ocupate si pune-le in positions_held.\n"
        "- current_position = ultima functie sau functia dominanta din CV.\n"
        "- recommended_role_for_candidate = rolul real potrivit candidatului, pe baza CV-ului. Nu copia automat postul cautat.\n"
        "- summary trebuie sa fie un rezumat HR curat. Nu include in summary textele: Risc supracalificare, Nepotrivire nivel, overqualification_risk, level_mismatch. Acestea exista separat in JSON.\n"
        "- Respecta must_have, nice_to_have, reject_if_missing, red_flags si max_score_rules.\n"
        "- Alege candidatul potrivit pentru postul cautat, nu candidatul cu cel mai impresionant CV.\n"
        "- Pentru roluri operationale, entry-level sau repetitive, penalizeaza supracalificarea.\n"
        "- Daca overqualification_risk este high, final_score maxim 40.\n"
        "- Daca lipseste o cerinta din reject_if_missing, final_score maxim 35.\n"
        "- Daca postul cere permis si permisul cerut lipseste, final_score maxim 30.\n\n"
        "SCORING:\n"
        "85-100 STRONG_YES, 70-84 YES, 55-69 MAYBE, 40-54 NO, sub 40 REJECT.\n\n"
        "Raspunde STRICT in JSON valid, fara markdown, folosind schema:\n" + schema + "\nCV:\n" + text[:10000]
    )


def analyze_cv_with_ai(text, target_job):
    if not os.getenv("OPENROUTER_API_KEY"):
        raise ValueError("Lipseste OPENROUTER_API_KEY din .env")
    model = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini")
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": build_prompt(text, target_job)}]
    )
    return response.choices[0].message.content


def parse_ai_response(content):
    if not content:
        return None
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", content, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return None
    return None


def clean_score(v):
    try:
        return max(0, min(100, int(v)))
    except Exception:
        return 0


def as_text(v):
    if v is None:
        return ""
    if isinstance(v, list):
        return ", ".join(str(x) for x in v)
    if isinstance(v, dict):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


def filename_to_name(file):
    clean = os.path.splitext(file)[0].replace("-"," ").replace("_"," ").replace("."," ")
    bad = {"cv","c","v","curriculum","vitae","resume","formal","2024","2025","pdf"}
    words = [w for w in clean.split() if w.lower() not in bad]
    return (" ".join(words).strip() or clean).title()


def normalize_recommendation(value, score):
    value = as_text(value).upper().strip()
    mapping = {"STRONG_YES":"HIRE","YES":"HIRE","MAYBE":"CONSIDER","NO":"REJECT","REJECT":"REJECT","HIRE":"HIRE","CONSIDER":"CONSIDER"}
    if value in mapping:
        return mapping[value]
    if score >= 70:
        return "HIRE"
    if score >= 45:
        return "CONSIDER"
    return "REJECT"


def recommendation_to_status(rec):
    return {"HIRE":"ADMIS","CONSIDER":"DE ANALIZAT","REJECT":"RESPINS"}.get(str(rec).upper().strip(), "NOU")


def apply_local_safety_rules(data, target_job, cv_text, profile):
    cv = normalize_text(cv_text)
    combined = normalize_text(" ".join([as_text(data.get(k,"")) for k in ["name","position","summary","strengths","skills"]]) + " " + cv[:3000])
    operational_domains = ["Depozit Logistica","Retail Financiar","Retail","Transport Soferie","Transport Livrari","Productie","Constructii","HoReCa","Facility","Securitate"]
    high_level = ["rector","profesor universitar","director general","ceo","antreprenor","consultant strategic","decan","academic"]
    if profile.get("domain") in operational_domains and any(w in combined for w in high_level):
        data["score"] = min(clean_score(data.get("score",0)), 35)
        data["recommendation"] = "REJECT"
        data["summary"] = set_summary_risk(data.get("summary",""), over_risk="high", level_risk="high")
    req = detect_required_license(target_job, profile)
    if req:
        r = req.lower()
        found = (f"permis {r}" in cv or f"categoria {r}" in cv or f"cat {r}" in cv or (req=="CE" and "tir" in cv))
        if not found:
            data["score"] = min(clean_score(data.get("score",0)), 30)
            data["recommendation"] = "REJECT"
            clean_visible = strip_risk_text(data.get("summary",""))
            data["summary"] = set_summary_risk(f"Lipseste confirmarea clara pentru permis categoria {req}. {clean_visible}", over_risk="high", level_risk="high")
    return data


def strip_risk_text(value):
    text = as_text(value).strip()
    text = re.sub(r"\bRisc\s+supracalificare\s*:\s*(low|medium|high|unknown|scazut|mediu|ridicat|necunoscut)\s*\.?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bNepotrivire\s+nivel\s*:\s*(low|medium|high|unknown|scazut|mediu|ridicat|necunoscut)\s*\.?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\boverqualification_risk\s*:\s*(low|medium|high|unknown)\s*\.?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\blevel_mismatch\s*:\s*(low|medium|high|unknown)\s*\.?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def set_summary_risk(summary, over_risk=None, level_risk=None):
    text = as_text(summary)
    over = over_risk or "unknown"
    level = level_risk or "unknown"

    current_over = re.search(r"Risc supracalificare:\s*(low|medium|high|unknown)", text, flags=re.IGNORECASE)
    current_level = re.search(r"Nepotrivire nivel:\s*(low|medium|high|unknown)", text, flags=re.IGNORECASE)

    if current_over:
        text = re.sub(r"Risc supracalificare:\s*(low|medium|high|unknown)", f"Risc supracalificare: {over}", text, flags=re.IGNORECASE)
    else:
        text = f"Risc supracalificare: {over}. " + text

    if current_level:
        text = re.sub(r"Nepotrivire nivel:\s*(low|medium|high|unknown)", f"Nepotrivire nivel: {level}", text, flags=re.IGNORECASE)
    else:
        parts = text.split(".", 1)
        if len(parts) == 2 and "Risc supracalificare:" in parts[0]:
            text = parts[0] + f". Nepotrivire nivel: {level}." + parts[1]
        else:
            text = f"Nepotrivire nivel: {level}. " + text

    return re.sub(r"\s+", " ", text).strip()


def first_value(*values):
    for value in values:
        text = as_text(value).strip()
        if text:
            return text
    return ""


def normalize_data(data, file, target_job, cv_text, profile):
    score = clean_score(data.get("final_score", data.get("fit_percentage", 0)))
    rec = normalize_recommendation(data.get("recommendation",""), score)
    name = data.get("candidate_name") or data.get("name") or data.get("full_name") or filename_to_name(file)

    over_risk = as_text(data.get("overqualification_risk", "unknown")).lower() or "unknown"
    level_risk = as_text(data.get("level_mismatch", "unknown")).lower() or "unknown"

    clean_ai_summary = strip_risk_text(data.get("summary", data.get("match_analysis", "")))
    if not clean_ai_summary:
        clean_ai_summary = "Rezumat indisponibil. Verifica manual CV-ul."

    summary = f"Risc supracalificare: {over_risk}. Nepotrivire nivel: {level_risk}. {clean_ai_summary}"

    recommended_role = first_value(
        data.get("recommended_role_for_candidate"),
        data.get("current_position"),
        data.get("target_position"),
        data.get("position")
    )

    normalized = {
        "name": name,
        "email": data.get("email",""),
        "phone": data.get("phone",""),
        "position": recommended_role,
        "years_experience": data.get("years_experience",""),
        "skills": data.get("transferable_skills", data.get("skills","")),
        "companies": data.get("companies",""),
        "score": score,
        "level": data.get("confidence_level", data.get("level","")),
        "strengths": data.get("strengths",""),
        "weaknesses": data.get("gaps", data.get("weaknesses","")),
        "summary": summary,
        "recommendation": rec
    }
    return apply_local_safety_rules(normalized, target_job, cv_text, profile)


def save_candidate_to_db(data, file, target_job):
    try:
        db = SessionLocal()
        candidate = Candidate(
            name=as_text(data.get("name","")),
            email=as_text(data.get("email","")),
            phone=as_text(data.get("phone","")),
            position=as_text(data.get("position","")),
            experience=as_text(data.get("years_experience","")),
            skills=as_text(data.get("skills","")),
            companies=as_text(data.get("companies","")),
            score=clean_score(data.get("score",0)),
            level=as_text(data.get("level","")),
            strengths=as_text(data.get("strengths","")),
            weaknesses=as_text(data.get("weaknesses","")),
            summary=as_text(data.get("summary","")),
            job_title=target_job,
            status=recommendation_to_status(data.get("recommendation","CONSIDER")),
            cv_file=file
        )
        db.add(candidate)
        db.commit()
        db.close()
        print("✓ SALVAT IN BAZA DE DATE:", data.get("name",""), "-", candidate.status)
    except Exception as e:
        print("EROARE LA SALVARE IN DB:", e)


# ─── PARALLEL TASK ────────────────────────────────────────────────────────────

def _process_single_cv(args):
    """
    Proceseaza un singur CV — folosit de ThreadPoolExecutor.
    Returneaza dict cu rezultatul sau eroarea.
    """
    file, target_job, profile = args

    # 1. Verifica cache
    cached = get_cached_analysis(file, target_job)
    if cached:
        logger.info(f"CACHE HIT: {file} pentru {target_job}")
        return {"file": file, "data": cached, "from_cache": True, "error": None}

    # 2. Citeste text din fisier
    path = os.path.join(UPLOAD_FOLDER, file)
    if not os.path.exists(path):
        return {"file": file, "data": None, "from_cache": False, "error": "Fisier negasit local"}

    text = extract_text_from_file(path)
    if not text:
        return {"file": file, "data": None, "from_cache": False, "error": "Nu s-a putut extrage text"}

    # 3. Analiza AI
    try:
        ai_response = analyze_cv_with_ai(text, target_job)
        data = parse_ai_response(ai_response)
        if not data:
            return {"file": file, "data": None, "from_cache": False, "error": "Raspuns AI invalid"}

        normalized = normalize_data(data, file, target_job, text, profile)

        # 4. Salveaza in cache pentru viitor
        save_analysis_to_cache(file, target_job, normalized)

        return {"file": file, "data": normalized, "from_cache": False, "error": None}

    except Exception as e:
        return {"file": file, "data": None, "from_cache": False, "error": str(e)}


# ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────

def process_cvs_for_job(target_job):
    target_job = target_job.strip()
    if not target_job:
        return {"ok": False, "message": "Nu ai completat postul.", "saved": 0}

    profile = match_job_profile(target_job) or build_fallback_profile(target_job)
    print("Profil job folosit:", profile.get("job_title"), "-", profile.get("domain"))

    # Colecteaza fisierele CV disponibile
    try:
        all_files = [
            f for f in os.listdir(UPLOAD_FOLDER)
            if f.lower().endswith((".pdf", ".docx", ".doc"))
        ]
    except FileNotFoundError:
        all_files = []

    if not all_files:
        return {"ok": False, "message": "Nu exista CV-uri in uploads. Incarca CV-uri mai intai.", "saved": 0}

    # Numara cate sunt din cache vs. noi
    cache_hits = 0
    fresh_analyses = 0
    saved = 0
    results = []

    # Procesare paralela — max 5 fire simultane (evita rate limiting API)
    tasks = [(file, target_job, profile) for file in all_files]

    print(f"\nProcesez {len(tasks)} CV-uri pentru postul: {target_job}")
    print("=" * 70)

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(_process_single_cv, task): task[0] for task in tasks}

        for future in as_completed(futures):
            result = future.result()
            file = result["file"]

            if result["error"]:
                print(f"✗ EROARE {file}: {result['error']}")
                continue

            data = result["data"]
            if not data:
                continue

            if result["from_cache"]:
                cache_hits += 1
                print(f"⚡ CACHE: {file} -> {data.get('name')} (scor: {data.get('score')})")
            else:
                fresh_analyses += 1
                print(f"✓ ANALIZAT: {file} -> {data.get('name')} (scor: {data.get('score')})")

            save_candidate_to_db(data, file, target_job)
            saved += 1
            results.append({
                "name": data.get("name"),
                "score": data.get("score"),
                "recommendation": data.get("recommendation")
            })

    results.sort(key=lambda x: x.get("score", 0), reverse=True)

    print("\n" + "=" * 70)
    print("SUMMARY:")
    print(f"  Din cache (instant): {cache_hits}")
    print(f"  Analizate acum:      {fresh_analyses}")
    print(f"  Total salvati:       {saved}")
    print("=" * 70)

    parts = []
    if fresh_analyses > 0:
        parts.append(f"{fresh_analyses} CV-uri analizate acum")
    if cache_hits > 0:
        parts.append(f"{cache_hits} preluate instant din cache")

    message = f"Analiza finalizata. {', '.join(parts)}. Total candidati: {saved}."

    return {
        "ok": True,
        "message": message,
        "saved": saved,
        "cache_hits": cache_hits,
        "fresh_analyses": fresh_analyses,
        "results": results
    }


def parse_all_cvs():
    target_job = input("Pentru ce post cauti candidat? ")
    process_cvs_for_job(target_job)


if __name__ == "__main__":
    parse_all_cvs()
