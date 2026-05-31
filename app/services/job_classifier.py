"""
Motor de clasificare si matching realist pentru NEXAS HR.
Nu inlocuieste AI-ul. Il controleaza.
AI extrage informatii. Acest modul decide eligibilitatea reala pe familie/subfamilie.
"""

import re
import unicodedata
from typing import Dict, Tuple

from app.services.job_taxonomy import JOB_TAXONOMY, RELATED_FAMILIES, STRICT_FAMILIES


def normalize_text(value: str) -> str:
    text = str(value or "").lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("ș", "s").replace("ţ", "t").replace("ț", "t").replace("ă", "a").replace("â", "a").replace("î", "i")
    text = re.sub(r"[^a-z0-9+#.\-/ ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _contains_phrase(text: str, phrase: str) -> bool:
    p = normalize_text(phrase)
    if not p:
        return False
    if len(p) <= 3 or any(ch in p for ch in "+#."):
        return p in text
    return re.search(r"(?<![a-z0-9])" + re.escape(p) + r"(?![a-z0-9])", text) is not None


def _score_family(text: str, family_key: str, spec: Dict) -> Tuple[int, str, int]:
    score = 0
    best_subfamily = "general"
    best_sub_score = 0

    for alias in spec.get("aliases", []):
        if _contains_phrase(text, alias):
            score += 4 if " " in alias else 2

    for sub_key, words in spec.get("subfamilies", {}).items():
        sub_score = 0
        for word in words:
            if _contains_phrase(text, word):
                sub_score += 5 if " " in word else 3
        if sub_score > best_sub_score:
            best_sub_score = sub_score
            best_subfamily = sub_key
        score += sub_score

    return score, best_subfamily, best_sub_score


def classify_job(job_title: str, extra_text: str = "") -> Dict:
    text = normalize_text(f"{job_title} {extra_text}")
    best = {"family": "UNKNOWN", "subfamily": "general", "score": 0, "confidence": 0.0}

    for family, spec in JOB_TAXONOMY.items():
        score, subfamily, sub_score = _score_family(text, family, spec)
        if score > best["score"]:
            best = {
                "family": family,
                "subfamily": subfamily,
                "score": score,
                "sub_score": sub_score,
                "confidence": min(0.98, score / 18),
                "label": spec.get("label", family),
            }

    return best


def classify_cv(cv_text: str, candidate_role: str = "") -> Dict:
    # Titlul candidatului conteaza mult, dar nu suficient singur.
    text = normalize_text(f"{candidate_role} {candidate_role} {candidate_role} {cv_text}")
    best = {"family": "UNKNOWN", "subfamily": "general", "score": 0, "confidence": 0.0}

    for family, spec in JOB_TAXONOMY.items():
        score, subfamily, sub_score = _score_family(text, family, spec)
        if score > best["score"]:
            best = {
                "family": family,
                "subfamily": subfamily,
                "score": score,
                "sub_score": sub_score,
                "confidence": min(0.98, score / 28),
                "label": spec.get("label", family),
            }

    return best


def has_must_have_signal(family: str, text: str) -> bool:
    spec = JOB_TAXONOMY.get(family, {})
    normalized = normalize_text(text)
    signals = spec.get("must_have_any", [])
    if not signals:
        return True
    return any(_contains_phrase(normalized, signal) for signal in signals)


def candidate_blob_for_matching(candidate) -> str:
    fields = [
        "name", "position", "job_title", "skills", "companies", "summary", "strengths", "weaknesses",
        "decision_reason", "missing_requirements", "manager_summary", "better_role_match",
    ]
    parts = []
    for field in fields:
        try:
            value = getattr(candidate, field, "")
        except Exception:
            value = ""
        if value:
            parts.append(str(value))
    return " ".join(parts)


def evaluate_realistic_fit(target_job: str, candidate_role: str = "", cv_text: str = "") -> Dict:
    """
    Returneaza verdict stabil pentru sortare si plafonare scor.

    fit_level:
    - direct: aceeasi familie si subfamilie apropiata
    - near: aceeasi familie, alta subfamilie sau familie foarte apropiata cu semnale clare
    - transferable: familie vecina, dar nu pentru top principal
    - wrong_domain: familie gresita sau lipsa semnal eliminatoriu
    """
    job = classify_job(target_job)
    cv = classify_cv(cv_text, candidate_role=candidate_role)

    job_family = job.get("family", "UNKNOWN")
    cv_family = cv.get("family", "UNKNOWN")
    job_sub = job.get("subfamily", "general")
    cv_sub = cv.get("subfamily", "general")

    all_text = f"{candidate_role} {cv_text}"
    job_known = job_family != "UNKNOWN" and job.get("score", 0) > 0
    cv_known = cv_family != "UNKNOWN" and cv.get("score", 0) > 0

    if not job_known:
        return {
            "fit_level": "near",
            "rank": 3,
            "score": 60,
            "score_floor": 0,
            "score_cap": 80,
            "job_family": job_family,
            "candidate_family": cv_family,
            "job_subfamily": job_sub,
            "candidate_subfamily": cv_sub,
            "reason": "Jobul nu a fost recunoscut clar in taxonomie. Se foloseste scoring AI controlat.",
        }

    # Pentru familii stricte, candidatul trebuie sa aiba semnale reale din acea familie.
    if job_family in STRICT_FAMILIES and not has_must_have_signal(job_family, all_text):
        return {
            "fit_level": "wrong_domain",
            "rank": 0,
            "score": 18,
            "score_floor": 0,
            "score_cap": 25,
            "job_family": job_family,
            "candidate_family": cv_family,
            "job_subfamily": job_sub,
            "candidate_subfamily": cv_sub,
            "reason": f"Lipsesc semnale eliminatorii pentru {JOB_TAXONOMY[job_family]['label']}.",
        }

    if cv_family == job_family:
        same_sub = cv_sub == job_sub or job_sub == "general" or cv_sub == "general"
        if same_sub:
            return {
                "fit_level": "direct",
                "rank": 4,
                "score": 86,
                "score_floor": 72,
                "score_cap": 96,
                "job_family": job_family,
                "candidate_family": cv_family,
                "job_subfamily": job_sub,
                "candidate_subfamily": cv_sub,
                "reason": f"Potrivire directa: {JOB_TAXONOMY[job_family]['label']} / {job_sub}.",
            }
        return {
            "fit_level": "near",
            "rank": 3,
            "score": 74,
            "score_floor": 62,
            "score_cap": 84,
            "job_family": job_family,
            "candidate_family": cv_family,
            "job_subfamily": job_sub,
            "candidate_subfamily": cv_sub,
            "reason": f"Aceeasi familie profesionala, subfamilie diferita: {cv_sub} vs {job_sub}.",
        }

    if cv_family in RELATED_FAMILIES.get(job_family, set()):
        # Pentru joburi stricte, familie vecina inseamna cel mult near/transferable, nu top peste direct match.
        if has_must_have_signal(job_family, all_text):
            return {
                "fit_level": "near",
                "rank": 2,
                "score": 65,
                "score_floor": 50,
                "score_cap": 72,
                "job_family": job_family,
                "candidate_family": cv_family,
                "job_subfamily": job_sub,
                "candidate_subfamily": cv_sub,
                "reason": f"Familie apropiata cu semnale pentru postul cautat: {cv_family} -> {job_family}.",
            }
        return {
            "fit_level": "transferable",
            "rank": 1,
            "score": 48,
            "score_floor": 0,
            "score_cap": 58,
            "job_family": job_family,
            "candidate_family": cv_family,
            "job_subfamily": job_sub,
            "candidate_subfamily": cv_sub,
            "reason": f"Familie apropiata, dar fara dovezi directe suficiente: {cv_family} -> {job_family}.",
        }

    # Daca CV-ul nu e clasificat, dar contine semnale obligatorii pentru job, il lasam ca near.
    if not cv_known and has_must_have_signal(job_family, all_text):
        return {
            "fit_level": "near",
            "rank": 2,
            "score": 62,
            "score_floor": 45,
            "score_cap": 75,
            "job_family": job_family,
            "candidate_family": cv_family,
            "job_subfamily": job_sub,
            "candidate_subfamily": cv_sub,
            "reason": f"CV neclasificat clar, dar are semnale pentru {JOB_TAXONOMY[job_family]['label']}.",
        }

    return {
        "fit_level": "wrong_domain",
        "rank": 0,
        "score": 12,
        "score_floor": 0,
        "score_cap": 25 if job_family in STRICT_FAMILIES else 35,
        "job_family": job_family,
        "candidate_family": cv_family,
        "job_subfamily": job_sub,
        "candidate_subfamily": cv_sub,
        "reason": f"Familie profesionala diferita: {cv_family} -> {job_family}.",
    }
