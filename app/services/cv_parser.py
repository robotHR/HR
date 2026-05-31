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

from app.services.cloudinary_service import upload_cv_to_cloudinary, check_cv_exists_on_cloudinary, stream_cv_from_cloudinary

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

UPLOAD_FOLDER = "app/uploads"


def sync_from_cloudinary():
    """
    Descarca din Cloudinary toate fisierele care nu exista local in app/uploads/.
    Apelat automat inainte de analiza pentru a asigura ca fisierele sunt disponibile.
    Returneaza numarul de fisiere sincronizate.
    """
    import cloudinary.api
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    try:
        results = cloudinary.api.resources(
            resource_type="raw",
            max_results=500
        )
        remote_files = results.get("resources", [])
    except Exception as e:
        logger.warning(f"Nu s-au putut lista fisierele din Cloudinary: {e}")
        return 0

    synced = 0
    for resource in remote_files:
        filename = resource.get("public_id", "")
        if not filename:
            continue

        ext = os.path.splitext(filename)[1].lower()
        if ext not in (".pdf", ".docx", ".doc"):
            continue

        local_path = os.path.join(UPLOAD_FOLDER, filename)

        if os.path.exists(local_path):
            continue

        buffer, mime = stream_cv_from_cloudinary(filename)
        if buffer:
            with open(local_path, "wb") as f:
                f.write(buffer.read())
            synced += 1
            logger.info(f"Sincronizat din Cloudinary: {filename}")

    if synced > 0:
        logger.info(f"Sincronizare Cloudinary completa: {synced} fisiere descarcate local")
    return synced

TEMP_FOLDER = "/tmp/nexas_hr"
os.makedirs(TEMP_FOLDER, exist_ok=True)

JOB_PROFILES_PATH = "config/job_profiles.json"
client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=os.getenv("OPENROUTER_API_KEY"))

# ─── CACHE FUNCTIONS ──────────────────────────────────────────────────────────

def _normalize_job_key(job_title):
    return normalize_text(job_title).strip()


def get_cached_analysis(cv_file, job_title):
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
    job_lower = normalize_text(target_job)

    operational_kw = [
        "operator", "muncitor", "lucrator", "manipulant", "depozitar", "stivuitorist",
        "curier", "casier", "ospatar", "vanzator", "agent vanzari", "curatenie",
        "ingrijitor", "paznic", "agent paza", "spalator", "sorter", "ambalator",
        "picker", "packer", "necalificat", "zilier", "ajutor",
    ]
    skilled_trade_kw = [
        "electrician", "sudor", "lacatus", "mecanic auto", "mecanic", "instalator",
        "zugrav", "zugrav vopsitor", "tamplar", "zidar", "faiantar", "montator",
        "frigotehnist", "automatist", "mentenanta", "tehnic",
    ]
    senior_kw = [
        "coordonator", "manager", "director", "analist", "inginer", "contabil",
        "economist", "auditor", "consultant", "specialist", "expert", "sef",
        "responsabil hr", "hr", "jurist", "avocat", "financiar", "controller",
        "programator", "developer", "it", "arhitect", "medic", "farmacist",
        # Roluri profesionale/administrative care cer studii superioare
        "consilier", "inspector", "referent", "judecator", "procuror", "notar",
        "diplomat", "ofiter", "parlamentar", "prefect", "primar",
        "sef birou", "sef serviciu", "functionar public",
    ]

    is_operational = any(kw in job_lower for kw in operational_kw)
    is_skilled_trade = any(kw in job_lower for kw in skilled_trade_kw)
    is_senior = any(kw in job_lower for kw in senior_kw)

    if is_operational:
        return {
            "job_title": target_job, "domain": "Operational", "level": "Entry",
            "requires_higher_education": False,
            "must_have": ["seriozitate", "punctualitate", "disponibilitate program"],
            "nice_to_have": ["experienta in domeniu similar", "permis auto"],
            "reject_if_missing": [],
            "overqualified_risk": ["licenta", "master", "doctorat", "inginer", "economist", "jurist", "studii superioare"],
            "red_flags": [],
            "max_score_rules": [
                "REGULA CRITICA: Daca candidatul are studii superioare (licenta/master/doctorat), final_score MAXIM 35 — supracalificat, nu va ramane.",
                "Daca candidatul nu are nicio experienta de munca, final_score maxim 58.",
                "Candidatul cu experienta in domenii similare (depozit/logistica/munca fizica/vanzari) poate lua 60-85.",
            ],
        }
    elif is_skilled_trade:
        return {
            "job_title": target_job, "domain": "Meserie calificata", "level": "Skilled",
            "requires_higher_education": False,
            "must_have": ["calificare sau experienta practica in meserie"],
            "nice_to_have": ["certificate autorizatii", "experienta recenta"],
            "reject_if_missing": ["experienta practica sau calificare in meserie"],
            "overqualified_risk": ["master", "doctorat"],
            "red_flags": ["fara experienta practica", "doar studii teoretice"],
            "max_score_rules": [
                "Fara experienta sau calificare practica in meserie, final_score maxim 38.",
                "Candidat cu master/doctorat fara experienta practica, final_score maxim 35.",
            ],
        }
    elif is_senior:
        return {
            "job_title": target_job, "domain": "Calificat/Senior/Profesional", "level": "Senior",
            "requires_higher_education": True,
            "must_have": ["studii superioare relevante pentru domeniu", "experienta minima in domeniu"],
            "nice_to_have": ["certificari", "limbi straine", "experienta de conducere"],
            "reject_if_missing": ["studii superioare in domeniu relevant", "experienta in domeniu"],
            "overqualified_risk": [],
            "red_flags": ["fara studii superioare", "fara experienta relevanta in domeniu", "background complet diferit"],
            "max_score_rules": [
                "REGULA CRITICA: Fara studii superioare in domeniu DIRECT relevant, final_score MAXIM 35.",
                "REGULA CRITICA: Fara NICIO experienta in domeniu sau domenii conexe, final_score MAXIM 38.",
                "REGULA CRITICA: Daca background-ul candidatului e complet diferit de domeniu (ex: vanzator/stivuitorist pentru rol juridic/parlamentar/financiar), final_score MAXIM 32.",
                "Experienta partiala sau tangentiala in domeniu: scor 39-55.",
                "Candidat cu studii relevante dar fara experienta: scor 40-55.",
            ],
        }
    else:
        return {
            "job_title": target_job, "domain": "General/Profesional", "level": "Middle",
            "requires_higher_education": False,
            "must_have": ["experienta in domeniu relevant sau conexe"],
            "nice_to_have": ["studii superioare", "certificari relevante"],
            "reject_if_missing": [],
            "overqualified_risk": [],
            "red_flags": ["lipsa oricarei experiente in domeniu sau domenii conexe"],
            "max_score_rules": [
                "Daca candidatul nu are NICIO experienta in domeniu sau domenii conexe, final_score MAXIM 38.",
                "Daca experienta este tangentiala (transferabilitate slaba), final_score maxim 48.",
                "Daca experienta este partial relevanta, scor 49-65.",
            ],
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


def build_prompt(text, target_job, job_requirements=None):
    profile = match_job_profile(target_job) or build_fallback_profile(target_job)
    profile_json = json.dumps(profile, ensure_ascii=False, indent=2)
    required_license = detect_required_license(target_job, profile)
    license_rule = f"Postul cere permis {required_license}. Nu cere alta categorie." if required_license else "Postul nu cere permis explicit."

    job_req_block = ""
    if job_requirements and job_requirements.strip():
        job_req_block = (
            "CERINTE REALE ALE POSTULUI (scrise de HR — PRIORITATE MAXIMA):\n"
            + job_requirements.strip()
            + "\n\nFOLOSESTE ACESTE CERINTE ca referinta principala pentru scoring. "
            "Daca CV-ul nu indeplineste cerintele obligatorii de mai sus, penalizeaza scorul conform regulilor de calibrare.\n\n"
        )

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
  "summary": "",
  "decision_reason": "",
  "missing_requirements": [],
  "must_verify_by_phone": [],
  "recommended_next_action": "CALL_NOW|CALL_LATER|KEEP_FOR_OTHER_ROLE|REJECT",
  "priority": "HIGH|MEDIUM|LOW",
  "manager_summary": "",
  "phone_call_script": "",
  "better_role_match": "",
  "reject_reason_internal": ""
}
"""
    return (
        "TU ESTI RECRUITER HR SENIOR CU 15+ ANI EXPERIENTA IN ROMANIA.\n\n"
        f"ANALIZEZI CV-UL PENTRU POSTUL: {target_job}\n\n"
        + job_req_block
        + "PROFIL JOB:\n" + profile_json + "\n\n"
        f"REGULA PERMIS: {license_rule}\n\n"
        "═══════════════════════════════════════════════════════════\n"
        "CALIBRARE NIVEL JOB — REGULI STRICTE DE SCORING\n"
        "═══════════════════════════════════════════════════════════\n"
        "Primul pas: determina nivelul jobului din titlu si cerinte.\n\n"
        "JOB OPERATIONAL/ENTRY (operator, muncitor, depozitar, casier, curier, vanzator, paznic, stivuitorist, etc.):\n"
        "  → Candidat cu studii superioare (licenta/master/doctorat) = SUPRACALIFICAT\n"
        "    final_score MAXIM 35, overqualification_risk=high\n"
        "    Motivul: nu va ramane, va pleca in 1-2 luni. Nu e potrivit.\n"
        "  → Candidat fara studii superioare + experienta in domenii similare = IDEAL\n"
        "    Poate lua scor 60-90 in functie de potrivire.\n"
        "  → Candidat fara nicio experienta dar fara studii superioare = ACCEPTABIL\n"
        "    Poate lua scor 45-60 (job accesibil, se poate instrui).\n\n"
        "JOB CALIFICAT/SENIOR (coordonator, manager, director, analist, inginer, contabil, HR,\n"
        "   financiar, IT, arhitect, medic, farmacist, consilier, inspector, referent,\n"
        "   judecator, procuror, notar, diplomat, functionar public, etc.):\n"
        "  → Candidat fara studii superioare in domeniu DIRECT relevant = SUBCALIFICAT\n"
        "    final_score MAXIM 35\n"
        "  → Candidat fara NICIO experienta in domeniu sau conexe = SUBCALIFICAT\n"
        "    final_score MAXIM 38\n"
        "  → Candidat cu background COMPLET DIFERIT (ex: vanzator/depozitar/sofer aplicand\n"
        "    pentru consilier parlamentar, inspector fiscal, judecator, contabil, inginer, etc.):\n"
        "    final_score MAXIM 32. Abilitatile de comunicare sau soft skills NU compenseaza\n"
        "    lipsa totala de educatie si experienta in domeniu.\n"
        "  → Candidat cu experienta partiala in domeniu: scor 39-55\n\n"
        "JOB MESERIE CALIFICATA (electrician, sudor, mecanic, instalator, etc.):\n"
        "  → Fara calificare/certificare practica = final_score MAXIM 38\n"
        "  → Cu calificare dar fara experienta recenta = scor 40-55\n\n"
        "REGULA CRITICA ANTI-INFLATIE SCOR:\n"
        "Soft skills (comunicare, adaptabilitate, seriozitate) NU pot compensa lipsa\n"
        "educatiei sau experientei de domeniu pentru roluri specializate.\n"
        "Un vanzator NU poate lua 60 pentru 'consilier parlamentar' indiferent cat\n"
        "de bun comunicator e — nu are pregatirea necesara.\n"
        "Scorul reflecta POTRIVIREA REALA, nu potentialul abstract.\n"
        "═══════════════════════════════════════════════════════════\n\n"
        "REGULI OBLIGATORII:\n"
        "- Extrage companiile/institutiile unde a lucrat candidatul si pune-le in companies. Nu inventa companii. Daca nu apar clar, lasa lista goala.\n"
        "- Extrage functiile ocupate si pune-le in positions_held.\n"
        "- current_position = ultima functie sau functia dominanta din CV.\n"
        "- recommended_role_for_candidate = jobul cel mai potrivit pentru acest candidat pe piata muncii din Romania, bazat STRICT pe educatia, experienta si skillurile din CV. IGNORA complet postul cautat. Gandeste independent: daca omul are experienta de contabil, scrie 'Contabil'. Daca are experienta de sofer, scrie 'Sofer'. Daca are facultate tehnica si experienta in mentenanta, scrie 'Tehnician mentenanta'. NU copia job_title. NU lasa gol. Scrie intotdeauna un job concret si real.\n"
        "- summary trebuie sa fie un rezumat HR curat. Nu include in summary textele: Risc supracalificare, Nepotrivire nivel, overqualification_risk, level_mismatch. Acestea exista separat in JSON.\n"
        "- Respecta must_have, nice_to_have, reject_if_missing, red_flags si max_score_rules.\n"
        "- Alege candidatul potrivit pentru postul cautat, nu candidatul cu cel mai impresionant CV.\n"
        "- Pentru roluri operationale, entry-level sau repetitive, penalizeaza supracalificarea.\n"
        "- Daca overqualification_risk este high, final_score maxim 40.\n"
        "- Daca lipseste o cerinta din reject_if_missing, final_score maxim 35.\n"
        "- Daca postul cere permis si permisul cerut lipseste, final_score maxim 30.\n\n"
        "REGULI PENTRU CAMPURILE NOI DE DECIZIE:\n"
        "- decision_reason = 1-2 propozitii clare de ce a primit acest scor. Concret, nu generic. Ex: 'Are 4 ani experienta directa in logistica si lucrul cu documente. Lipseste confirmarea pentru program in ture.'\n"
        "- missing_requirements = lista cu ce lipseste concret din CV, formulat specific. Ex: ['Nu apare permis categoria B', 'Nu apare experienta cu Excel', 'Nu apare disponibilitatea pentru ture', 'Nu apare salariul dorit']. Daca nu lipseste nimic important, lasa lista goala.\n"
        "- must_verify_by_phone = 3-5 intrebari scurte, concrete, de pus la telefon. Ex: ['Ai permis categoria B valabil?', 'Poti incepe in urmatoarele 2 saptamani?', 'Care este salariul net dorit?', 'Ai disponibilitate pentru program in ture?']\n"
        "- recommended_next_action = CALL_NOW daca scor >= 75, CALL_LATER daca scor 55-74, KEEP_FOR_OTHER_ROLE daca scor < 55 dar candidatul e bun pe alt rol, REJECT daca nu are potential real.\n"
        "- priority = HIGH daca scor >= 80, MEDIUM daca scor 60-79, LOW daca scor < 60.\n"
        "- manager_summary = maxim 3 randuri pentru manager. Fara jargon HR. Ex: 'Candidat recomandat pentru interviu. Are 4 ani experienta relevanta si indeplineste cerintele principale. Trebuie verificata disponibilitatea pentru program si salariul dorit.'\n"
        "- phone_call_script = text gata de folosit la telefon, in romani, natural, profesional. Include salut, motiv apel, 3-4 intrebari cheie si inchidere. Maxim 8 randuri.\n"
        "- better_role_match = daca candidatul ar fi mai potrivit pentru alt rol decat cel cautat, scrie acel rol. Altfel lasa gol.\n"
        "- reject_reason_internal = motiv scurt de respingere pentru uz intern HR. Nu se trimite candidatului. Scrie doar daca recommendation este NO sau REJECT.\n"
        "- interview_questions = 4-6 intrebari specifice pentru interviu, bazate pe CV-ul acestui candidat.\n\n"
        "SCORING:\n"
        "85-100 STRONG_YES, 70-84 YES, 55-69 MAYBE, 40-54 NO, sub 40 REJECT.\n\n"
        "Raspunde STRICT in JSON valid, fara markdown, folosind schema:\n" + schema + "\nCV:\n" + text[:15000]
    )


def analyze_cv_with_ai(text, target_job, job_requirements=None):
    if not os.getenv("OPENROUTER_API_KEY"):
        raise ValueError("Lipseste OPENROUTER_API_KEY din .env")
    model = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini")
    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": build_prompt(text, target_job, job_requirements)}]
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


def _compute_priority(score, data):
    """Calculeaza priority local ca fallback daca AI nu returneaza corect."""
    ai_priority = as_text(data.get("priority", "")).upper().strip()
    if ai_priority in ("HIGH", "MEDIUM", "LOW"):
        return ai_priority
    if score >= 80:
        return "HIGH"
    if score >= 60:
        return "MEDIUM"
    return "LOW"


def _compute_next_action(score, data):
    """Calculeaza recommended_next_action local ca fallback."""
    ai_action = as_text(data.get("recommended_next_action", "")).upper().strip()
    if ai_action in ("CALL_NOW", "CALL_LATER", "KEEP_FOR_OTHER_ROLE", "REJECT"):
        return ai_action
    if score >= 75:
        return "CALL_NOW"
    if score >= 55:
        return "CALL_LATER"
    better = as_text(data.get("better_role_match", "")).strip()
    if better:
        return "KEEP_FOR_OTHER_ROLE"
    return "REJECT"


def apply_local_safety_rules(data, target_job, cv_text, profile):
    cv = normalize_text(cv_text)
    combined = normalize_text(" ".join([as_text(data.get(k,"")) for k in ["name","position","summary","strengths","skills"]]) + " " + cv[:3000])
    operational_domains = ["Depozit Logistica","Retail Financiar","Retail","Transport Soferie","Transport Livrari","Productie","Constructii","HoReCa","Facility","Securitate"]
    high_level = ["rector","profesor universitar","director general","ceo","antreprenor","consultant strategic","decan","academic"]
    if profile.get("domain") in operational_domains and any(w in combined for w in high_level):
        data["score"] = min(clean_score(data.get("score",0)), 35)
        data["recommendation"] = "REJECT"
        data["summary"] = set_summary_risk(data.get("summary",""), over_risk="high", level_risk="high")
        data["priority"] = "LOW"
        data["recommended_next_action"] = "REJECT"
        data["reject_reason_internal"] = "Supracalificare ridicata pentru rol operational."
    req = detect_required_license(target_job, profile)
    if req:
        r = req.lower()
        found = (f"permis {r}" in cv or f"categoria {r}" in cv or f"cat {r}" in cv or (req=="CE" and "tir" in cv))
        if not found:
            data["score"] = min(clean_score(data.get("score",0)), 30)
            data["recommendation"] = "REJECT"
            clean_visible = strip_risk_text(data.get("summary",""))
            data["summary"] = set_summary_risk(f"Lipseste confirmarea clara pentru permis categoria {req}. {clean_visible}", over_risk="high", level_risk="high")
            data["priority"] = "LOW"
            data["recommended_next_action"] = "REJECT"
            missing = as_text(data.get("missing_requirements", ""))
            permis_item = f"Nu apare permis categoria {req} in CV"
            if permis_item not in missing:
                existing = [x.strip() for x in missing.split(",") if x.strip()] if missing else []
                existing.insert(0, permis_item)
                data["missing_requirements"] = existing
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

    priority = _compute_priority(score, data)
    recommended_next_action = _compute_next_action(score, data)

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
        "recommendation": rec,
        # ── Campuri noi decizie ──
        "decision_reason":         as_text(data.get("decision_reason", "")),
        "missing_requirements":    as_text(data.get("missing_requirements", [])),
        "must_verify_by_phone":    as_text(data.get("must_verify_by_phone", [])),
        "recommended_next_action": recommended_next_action,
        "priority":                priority,
        "manager_summary":         as_text(data.get("manager_summary", "")),
        "phone_call_script":       as_text(data.get("phone_call_script", "")),
        "interview_questions":     as_text(data.get("interview_questions", [])),
        "better_role_match":       as_text(data.get("better_role_match", "")),
        "reject_reason_internal":  as_text(data.get("reject_reason_internal", "")),
    }
    return apply_local_safety_rules(normalized, target_job, cv_text, profile)


def save_candidate_to_db(data, file, target_job):
    try:
        db = SessionLocal()
        status = recommendation_to_status(data.get("recommendation", "CONSIDER"))
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
            status=status,
            cv_file=file,
            # ── Campuri noi ──
            decision_reason=as_text(data.get("decision_reason","")),
            missing_requirements=as_text(data.get("missing_requirements","")),
            must_verify_by_phone=as_text(data.get("must_verify_by_phone","")),
            recommended_next_action=as_text(data.get("recommended_next_action","")),
            priority=as_text(data.get("priority","")),
            manager_summary=as_text(data.get("manager_summary","")),
            phone_call_script=as_text(data.get("phone_call_script","")),
            interview_questions=as_text(data.get("interview_questions","")),
            better_role_match=as_text(data.get("better_role_match","")),
            reject_reason_internal=as_text(data.get("reject_reason_internal","")),
        )
        db.add(candidate)
        db.commit()
        candidate_id = candidate.id
        db.close()
        print("✓ SALVAT IN BAZA DE DATE:", data.get("name",""), "-", status)
        return candidate_id
    except Exception as e:
        print("EROARE LA SALVARE IN DB:", e)
        try:
            db.rollback()
            db.close()
        except Exception:
            pass
        return None


# ─── PARALLEL TASK ────────────────────────────────────────────────────────────

def _process_single_cv(args):
    file, target_job, profile, job_requirements = args

    cached = get_cached_analysis(file, target_job)
    if cached:
        logger.info(f"CACHE HIT: {file} pentru {target_job}")
        return {"file": file, "data": cached, "from_cache": True, "error": None}

    path = os.path.join(UPLOAD_FOLDER, file)
    if not os.path.exists(path):
        return {"file": file, "data": None, "from_cache": False, "error": "Fisier negasit local"}

    text = extract_text_from_file(path)
    if not text:
        return {"file": file, "data": None, "from_cache": False, "error": "Nu s-a putut extrage text"}

    try:
        ai_response = analyze_cv_with_ai(text, target_job, job_requirements)
        data = parse_ai_response(ai_response)
        if not data:
            return {"file": file, "data": None, "from_cache": False, "error": "Raspuns AI invalid"}

        normalized = normalize_data(data, file, target_job, text, profile)
        save_analysis_to_cache(file, target_job, normalized)

        return {"file": file, "data": normalized, "from_cache": False, "error": None}

    except Exception as e:
        return {"file": file, "data": None, "from_cache": False, "error": str(e)}


# ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────

def process_cvs_for_job(target_job, job_requirements=None):
    target_job = target_job.strip()
    if not target_job:
        return {"ok": False, "message": "Nu ai completat postul.", "saved": 0}

    profile = match_job_profile(target_job) or build_fallback_profile(target_job)
    print("Profil job folosit:", profile.get("job_title"), "-", profile.get("domain"))
    if job_requirements:
        print("Cerinte post din platforma: DA")

    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    synced = sync_from_cloudinary()
    if synced > 0:
        print(f"Sincronizare Cloudinary: {synced} fisiere descarcate local")

    try:
        all_files = [
            f for f in os.listdir(UPLOAD_FOLDER)
            if f.lower().endswith((".pdf", ".docx", ".doc"))
        ]
    except FileNotFoundError:
        all_files = []

    if not all_files:
        return {"ok": False, "message": "Nu exista CV-uri. Verifica Cloudinary sau incarca CV-uri.", "saved": 0}

    cache_hits = 0
    fresh_analyses = 0
    saved = 0
    results = []

    tasks = [(file, target_job, profile, job_requirements) for file in all_files]

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
                "recommendation": data.get("recommendation"),
                "priority": data.get("priority"),
                "recommended_next_action": data.get("recommended_next_action"),
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
