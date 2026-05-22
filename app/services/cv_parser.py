import os, re, json, logging
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

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)
UPLOAD_FOLDER = "app/uploads"
JOB_PROFILES_PATH = "config/job_profiles.json"
client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=os.getenv("OPENROUTER_API_KEY"))

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
        logger.error("Lipseste python-docx. Ruleaza: pip install python-docx")
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
        logger.error("Lipseste pywin32. Ruleaza: pip install pywin32")
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


def normalize_text(value):
    value = str(value or "").lower()
    for a,b in [("ș","s"),("ş","s"),("ă","a"),("â","a"),("î","i"),("ț","t"),("ţ","t"),("."," "),("-"," "),("_"," "),("/"," ")]:
        value = value.replace(a,b)
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
        "- Daca postul cautat este Senior Logistics Coordinator, dar CV-ul este de Rector, recommended_role_for_candidate trebuie sa fie Rector / Coordonator academic / Management educational, nu Senior Logistics Coordinator.\n"
        "- summary trebuie sa fie un rezumat HR curat. Nu include in summary textele: Risc supracalificare, Nepotrivire nivel, overqualification_risk, level_mismatch. Acestea exista separat in JSON.\n"
        "- Respecta must_have, nice_to_have, reject_if_missing, red_flags si max_score_rules.\n"
        "- Alege candidatul potrivit pentru postul cautat, nu candidatul cu cel mai impresionant CV.\n"
        "- Pentru roluri operationale, entry-level sau repetitive, penalizeaza supracalificarea.\n"
        "- Rector, profesor universitar, director general, CEO, antreprenor sau consultant strategic pentru lucrator depozit, casier, lucrator comercial, sofer livrari = risc mare de supracalificare.\n"
        "- Daca overqualification_risk este high, final_score maxim 40.\n"
        "- Daca lipseste o cerinta din reject_if_missing, final_score maxim 35.\n"
        "- Daca postul cere permis si permisul cerut lipseste, final_score maxim 30.\n"
        "- Nu confunda operator masini cu operator calculator.\n"
        "- Nu confunda logistica office cu manipulare marfa.\n"
        "- Nu confunda sofer categoria B cu sofer categoria C.\n\n"
        "SCORING:\n"
        "85-100 STRONG_YES, 70-84 YES, 55-69 MAYBE, 40-54 NO, sub 40 REJECT.\n\n"
        "Raspunde STRICT in JSON valid, fara markdown, folosind schema:\n" + schema + "\nCV:\n" + text[:10000]
    )

def analyze_cv_with_ai(text, target_job):
    if not os.getenv("OPENROUTER_API_KEY"):
        raise ValueError("Lipseste OPENROUTER_API_KEY din .env")
    model = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini")
    response = client.chat.completions.create(model=model, messages=[{"role":"user","content":build_prompt(text,target_job)}])
    return response.choices[0].message.content

def parse_ai_response(content):
    if not content: return None
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", content, re.DOTALL)
        if m:
            try: return json.loads(m.group(0))
            except Exception: return None
    return None

def clean_score(v):
    try: return max(0, min(100, int(v)))
    except Exception: return 0

def as_text(v):
    if v is None: return ""
    if isinstance(v, list): return ", ".join(str(x) for x in v)
    if isinstance(v, dict): return json.dumps(v, ensure_ascii=False)
    return str(v)

def filename_to_name(file):
    clean = os.path.splitext(file)[0].replace("-"," ").replace("_"," ").replace("."," ")
    bad = {"cv","c","v","curriculum","vitae","resume","formal","2024","2025","pdf"}
    words = [w for w in clean.split() if w.lower() not in bad]
    return (" ".join(words).strip() or clean).title()

def normalize_recommendation(value, score):
    value = as_text(value).upper().strip()
    mapping = {"STRONG_YES":"HIRE","YES":"HIRE","MAYBE":"CONSIDER","NO":"REJECT","REJECT":"REJECT","HIRE":"HIRE","CONSIDER":"CONSIDER"}
    if value in mapping: return mapping[value]
    if score >= 70: return "HIRE"
    if score >= 45: return "CONSIDER"
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
        db.add(candidate); db.commit(); db.close()
        print("✓ SALVAT IN BAZA DE DATE:", data.get("name",""), "-", candidate.status)
    except Exception as e:
        print("EROARE LA SALVARE IN DB:", e)

def process_cvs_for_job(target_job):
    target_job = target_job.strip()
    if not target_job:
        return {"ok":False,"message":"Nu ai completat postul.","saved":0}
    profile = match_job_profile(target_job) or build_fallback_profile(target_job)
    print("Profil job folosit:", profile.get("job_title"), "-", profile.get("domain"))
    saved, results = 0, []
    for file in os.listdir(UPLOAD_FOLDER):
        if not file.lower().endswith((".pdf", ".docx", ".doc")):
            continue

        path = os.path.join(UPLOAD_FOLDER, file)
        print("\n" + "="*70); print(f"Processing: {file}"); print("="*70)

        text = extract_text_from_file(path)
        if not text:
            print("Nu am putut extrage text din:", file); continue
        ai_response = analyze_cv_with_ai(text, target_job)
        data = parse_ai_response(ai_response)
        if not data:
            print("Nu am putut interpreta raspunsul AI pentru:", file); print(ai_response); continue
        normalized = normalize_data(data, file, target_job, text, profile)
        print("Nume:", normalized.get("name")); print("Scor:", normalized.get("score")); print("Recomandare:", normalized.get("recommendation"))
        save_candidate_to_db(normalized, file, target_job)
        saved += 1
        results.append({"name":normalized.get("name"),"score":normalized.get("score"),"recommendation":normalized.get("recommendation")})
    results.sort(key=lambda x:x.get("score",0), reverse=True)
    print("\n" + "="*70); print("SUMMARY - CANDIDATI SORTATI DESCRESCATOR DUPA SCOR:"); print("="*70)
    for i,item in enumerate(results,1):
        print(f"{i}. {item.get('name')} - Score: {item.get('score')}/100 - {item.get('recommendation')}")
    return {"ok":True,"message":f"Analiza finalizata. Candidati salvati: {saved}","saved":saved,"results":results}

def parse_all_cvs():
    target_job = input("Pentru ce post cauti candidat? ")
    process_cvs_for_job(target_job)

if __name__ == "__main__":
    parse_all_cvs()