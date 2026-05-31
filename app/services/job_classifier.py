"""
Motor determinist de clasificare si filtrare HR.

Acest fisier tine scorul AI sub control. AI-ul poate descrie CV-ul,
dar eligibilitatea reala o decide acest motor pe familii de joburi.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from app.services.job_taxonomy import JOB_FAMILIES, ADJACENT_FAMILIES


def normalize_text(value: object) -> str:
    text = str(value or "").lower()
    replacements = {
        "ș": "s", "ş": "s", "ă": "a", "â": "a", "î": "i", "ț": "t", "ţ": "t",
        "-": " ", "_": " ", "/": " ", ".": " ", ",": " ", ";": " ", ":": " ", "|": " ", "\n": " "
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return " ".join(text.split())


def _contains_phrase(text: str, phrase: str) -> bool:
    phrase_n = normalize_text(phrase)
    if not phrase_n:
        return False
    return re.search(r"(^|\s)" + re.escape(phrase_n) + r"($|\s)", text) is not None


@dataclass
class Classification:
    family: str = "UNKNOWN"
    subfamily: str = "unknown"
    confidence: int = 0
    matched_terms: Tuple[str, ...] = ()

    @property
    def is_known(self) -> bool:
        return self.family != "UNKNOWN" and self.confidence > 0


@dataclass
class MatchDecision:
    include: bool
    relation: str
    score: int
    job_family: str
    job_subfamily: str
    candidate_family: str
    candidate_subfamily: str
    reason: str


def _score_family(text: str, title_hint: str = "") -> Dict[str, Tuple[int, str, List[str]]]:
    text_n = normalize_text(text)
    title_n = normalize_text(title_hint)
    scores: Dict[str, Tuple[int, str, List[str]]] = {}

    for family, config in JOB_FAMILIES.items():
        best_score = 0
        best_sub = "unknown"
        terms: List[str] = []
        for subfamily, aliases in config.get("subfamilies", {}).items():
            sub_score = 0
            sub_terms: List[str] = []
            for alias in aliases:
                alias_n = normalize_text(alias)
                if not alias_n:
                    continue
                if _contains_phrase(title_n, alias_n):
                    sub_score += 90 + min(len(alias_n.split()) * 4, 20)
                    sub_terms.append(alias)
                elif _contains_phrase(text_n, alias_n):
                    sub_score += 35 + min(len(alias_n.split()) * 3, 15)
                    sub_terms.append(alias)
            for signal in config.get("hard_signals", []):
                signal_n = normalize_text(signal)
                if _contains_phrase(text_n, signal_n):
                    sub_score += 8
                    sub_terms.append(signal)
            if sub_score > best_score:
                best_score = sub_score
                best_sub = subfamily
                terms = sub_terms[:8]
        if best_score > 0:
            scores[family] = (best_score, best_sub, terms)
    return scores


def classify_job(job_title: str) -> Classification:
    scores = _score_family(job_title, title_hint=job_title)
    if not scores:
        return Classification()
    family, (score, sub, terms) = max(scores.items(), key=lambda item: item[1][0])
    confidence = min(100, score)
    return Classification(family=family, subfamily=sub, confidence=confidence, matched_terms=tuple(terms))


def classify_candidate(candidate_text: str, role_hint: str = "") -> Classification:
    scores = _score_family(candidate_text, title_hint=role_hint)
    if not scores:
        return Classification()
    family, (score, sub, terms) = max(scores.items(), key=lambda item: item[1][0])
    confidence = min(100, score)
    return Classification(family=family, subfamily=sub, confidence=confidence, matched_terms=tuple(terms))


def _same_family_relation(job: Classification, cand: Classification) -> str:
    if job.subfamily == cand.subfamily:
        return "DIRECT"
    return "SAME_FAMILY"


def is_adjacent(job_family: str, candidate_family: str) -> bool:
    return candidate_family in ADJACENT_FAMILIES.get(job_family, set())


def evaluate_match(target_job: str, candidate_text: str, candidate_role: str = "", ai_score: int = 0, strict: bool = True) -> MatchDecision:
    job = classify_job(target_job)
    cand = classify_candidate(candidate_text, role_hint=candidate_role)
    ai = max(0, min(int(ai_score or 0), 100))

    if not job.is_known:
        # Cand nu recunoastem jobul, nu ascundem tot. Dar nu avem voie sa inventam compatibilitate mare.
        score = min(ai, 55)
        return MatchDecision(True, "UNKNOWN_JOB", score, job.family, job.subfamily, cand.family, cand.subfamily, "Jobul nu a fost recunoscut in taxonomie.")

    if not cand.is_known:
        return MatchDecision(False if strict else True, "UNKNOWN_CANDIDATE", min(ai, 25), job.family, job.subfamily, cand.family, cand.subfamily, "CV-ul nu are semnale clare pentru nicio familie profesionala.")

    if job.family == cand.family:
        relation = _same_family_relation(job, cand)
        if relation == "DIRECT":
            score = max(75, min(96, 78 + ai // 5 + cand.confidence // 12))
            return MatchDecision(True, relation, score, job.family, job.subfamily, cand.family, cand.subfamily, "Familie si subfamilie potrivite.")

        # Aceeasi familie, dar subfamilie diferita. Nu facem top fals.
        # Ex: bucatar la ospatar, IT QA la developer.
        score = max(45, min(68, 42 + ai // 6 + cand.confidence // 15))
        include = not strict or score >= 60
        return MatchDecision(include, relation, score, job.family, job.subfamily, cand.family, cand.subfamily, "Aceeasi familie, dar subfamilie diferita.")

    if is_adjacent(job.family, cand.family):
        score = min(45, max(25, 20 + ai // 8 + cand.confidence // 20))
        return MatchDecision(False if strict else True, "ADJACENT", score, job.family, job.subfamily, cand.family, cand.subfamily, "Familie apropiata, dar nu potrivire directa.")

    return MatchDecision(False if strict else True, "WRONG_FAMILY", min(ai, 18), job.family, job.subfamily, cand.family, cand.subfamily, "Familie profesionala diferita.")


def candidate_text_from_record(candidate) -> str:
    parts = [
        getattr(candidate, "position", ""),
        getattr(candidate, "job_title", ""),
        getattr(candidate, "skills", ""),
        getattr(candidate, "companies", ""),
        getattr(candidate, "experience", ""),
        getattr(candidate, "strengths", ""),
        getattr(candidate, "weaknesses", ""),
        getattr(candidate, "summary", ""),
        getattr(candidate, "cv_file", ""),
    ]
    return " ".join(str(p or "") for p in parts)


def evaluate_candidate_record(target_job: str, candidate, strict: bool = True) -> MatchDecision:
    return evaluate_match(
        target_job=target_job,
        candidate_text=candidate_text_from_record(candidate),
        candidate_role=getattr(candidate, "position", "") or "",
        ai_score=getattr(candidate, "score", 0) or 0,
        strict=strict,
    )


def recalibrate_ai_result(data: dict, target_job: str, cv_text: str, ai_score: int) -> dict:
    role = str(data.get("position") or data.get("recommended_role_for_candidate") or data.get("current_position") or "")
    decision = evaluate_match(target_job, cv_text, candidate_role=role, ai_score=ai_score, strict=True)

    data["score"] = decision.score
    data["candidate_family"] = decision.candidate_family
    data["job_family"] = decision.job_family
    data["candidate_subfamily"] = decision.candidate_subfamily
    data["job_subfamily"] = decision.job_subfamily
    data["match_relation"] = decision.relation
    data["candidate_cluster"] = decision.candidate_family
    data["job_cluster"] = decision.job_family

    if decision.relation in ("DIRECT",):
        data["recommendation"] = "STRONG_YES" if decision.score >= 85 else "YES"
        data["priority"] = "HIGH" if decision.score >= 80 else "MEDIUM"
        data["recommended_next_action"] = "CALL_NOW"
    elif decision.relation == "SAME_FAMILY":
        data["recommendation"] = "MAYBE"
        data["priority"] = "MEDIUM"
        data["recommended_next_action"] = "CALL_LATER"
    else:
        data["recommendation"] = "REJECT"
        data["priority"] = "LOW"
        data["recommended_next_action"] = "REJECT"
        data["reject_reason_internal"] = decision.reason

    base_summary = str(data.get("summary") or "").strip()
    family_line = (
        f"Potrivire familie: {decision.relation}. "
        f"Job={decision.job_family}/{decision.job_subfamily}. "
        f"CV={decision.candidate_family}/{decision.candidate_subfamily}. "
        f"{decision.reason}"
    )
    data["summary"] = (family_line + (" " + base_summary if base_summary else "")).strip()
    return data
