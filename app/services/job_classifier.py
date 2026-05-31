"""
Motor deterministic de clasificare HR — formula 50/30/20.

Familie  = 50 puncte
Subfamilie = 30 puncte
Specializare = 20 puncte
Total maxim = 100

Familie gresita = 0-5 puncte, candidatul NU apare in top.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from app.services.job_taxonomy import JOB_FAMILIES, ADJACENT_FAMILIES

logger = logging.getLogger("nexas.classifier")


def _normalize(value: object) -> str:
    text = str(value or "").lower()
    for src, dst in {
        "ș": "s", "ş": "s", "ă": "a", "â": "a", "î": "i", "ț": "t", "ţ": "t",
        "-": " ", "_": " ", "/": " ", ".": " ", ",": " ", ";": " ", ":": " ",
        "|": " ", "\n": " ", "\r": " ",
    }.items():
        text = text.replace(src, dst)
    return " ".join(text.split())


def _has_phrase(text: str, phrase: str) -> bool:
    p = _normalize(phrase)
    if not p:
        return False
    return re.search(r"(^|\s)" + re.escape(p) + r"($|\s)", text) is not None


# ─── CLASIFICARE FAMILIE ──────────────────────────────────────────────────────

@dataclass
class Classification:
    family: str = "UNKNOWN"
    subfamily: str = "unknown"
    specialization: str = "unknown"
    confidence: int = 0
    matched_terms: Tuple[str, ...] = ()

    @property
    def is_known(self) -> bool:
        return self.family != "UNKNOWN" and self.confidence > 0


def _classify(text: str, title_hint: str = "") -> Classification:
    """
    Clasifica un text in familia + subfamilia + specializarea potrivita.
    title_hint primeste bonus suplimentar (e mai precis decat textul full).
    """
    text_n = _normalize(text)
    title_n = _normalize(title_hint)

    best_family = "UNKNOWN"
    best_subfamily = "unknown"
    best_spec = "unknown"
    best_score = 0
    best_terms: List[str] = []

    for family, config in JOB_FAMILIES.items():
        subfamilies = config.get("subfamilies", {})
        specializations = config.get("specializations", {})
        hard_signals = config.get("hard_signals", [])

        for subfamily, aliases in subfamilies.items():
            sub_score = 0
            sub_terms: List[str] = []

            for alias in aliases:
                alias_n = _normalize(alias)
                if not alias_n:
                    continue
                if _has_phrase(title_n, alias_n):
                    # titlu exact — scor mare
                    sub_score += 80 + min(len(alias_n.split()) * 5, 20)
                    sub_terms.append(alias)
                elif _has_phrase(text_n, alias_n):
                    sub_score += 30 + min(len(alias_n.split()) * 3, 12)
                    sub_terms.append(alias)

            # hard_signals ca semnal secundar
            for signal in hard_signals:
                signal_n = _normalize(signal)
                if _has_phrase(text_n, signal_n):
                    sub_score += 5
                    sub_terms.append(signal)

            # specializari ca bonus
            found_spec = "unknown"
            for spec_name, spec_terms in specializations.get(subfamily, {}).items() if isinstance(specializations.get(subfamily), dict) else []:
                pass
            spec_list = specializations.get(subfamily, [])
            for spec in spec_list:
                spec_n = _normalize(spec)
                if _has_phrase(text_n, spec_n) or _has_phrase(title_n, spec_n):
                    sub_score += 3
                    found_spec = spec
                    break

            if sub_score > best_score:
                best_score = sub_score
                best_family = family
                best_subfamily = subfamily
                best_spec = found_spec
                best_terms = sub_terms[:8]

    confidence = min(100, best_score)
    return Classification(
        family=best_family,
        subfamily=best_subfamily,
        specialization=best_spec,
        confidence=confidence,
        matched_terms=tuple(best_terms),
    )


def classify_job(job_title: str) -> Classification:
    return _classify(job_title, title_hint=job_title)


def classify_candidate(candidate_text: str, role_hint: str = "") -> Classification:
    return _classify(candidate_text, title_hint=role_hint)


# ─── DECIZIE DE MATCHING ─────────────────────────────────────────────────────

@dataclass
class MatchDecision:
    include: bool
    relation: str          # DIRECT | SAME_FAMILY | ADJACENT | WRONG_FAMILY | UNKNOWN_JOB | UNKNOWN_CANDIDATE
    score: int             # 0-100
    family_score: int      # 0-50
    subfamily_score: int   # 0-30
    spec_score: int        # 0-20
    job_family: str
    job_subfamily: str
    candidate_family: str
    candidate_subfamily: str
    reason: str


def _is_adjacent(job_family: str, cand_family: str) -> bool:
    return cand_family in ADJACENT_FAMILIES.get(job_family, set())


def evaluate_match(
    target_job: str,
    candidate_text: str,
    candidate_role: str = "",
    ai_score: int = 0,
    strict: bool = True,
) -> MatchDecision:
    job = classify_job(target_job)
    cand = classify_candidate(candidate_text, role_hint=candidate_role)
    ai = max(0, min(int(ai_score or 0), 100))

    # ── Job necunoscut ───────────────────────────────────────────────────────
    if not job.is_known:
        score = min(ai, 55)
        _log(job, cand, 0, 0, 0, score, "UNKNOWN_JOB")
        return MatchDecision(
            include=True, relation="UNKNOWN_JOB", score=score,
            family_score=0, subfamily_score=0, spec_score=0,
            job_family=job.family, job_subfamily=job.subfamily,
            candidate_family=cand.family, candidate_subfamily=cand.subfamily,
            reason="Jobul nu a fost recunoscut in taxonomie.",
        )

    # ── Candidat necunoscut ──────────────────────────────────────────────────
    if not cand.is_known:
        score = min(ai, 20)
        _log(job, cand, 0, 0, 0, score, "UNKNOWN_CANDIDATE")
        return MatchDecision(
            include=False if strict else True,
            relation="UNKNOWN_CANDIDATE", score=score,
            family_score=0, subfamily_score=0, spec_score=0,
            job_family=job.family, job_subfamily=job.subfamily,
            candidate_family=cand.family, candidate_subfamily=cand.subfamily,
            reason="CV-ul nu are semnale clare pentru nicio familie.",
        )

    # ── Aceeasi familie ──────────────────────────────────────────────────────
    if job.family == cand.family:
        family_pts = 50  # 50/50

        if job.subfamily == cand.subfamily:
            # DIRECT — aceeasi subfamilie
            subfamily_pts = 30  # 30/30
            spec_pts = 20 if (cand.specialization != "unknown" and cand.specialization == job.family) else min(20, ai // 5)
            spec_pts = min(20, spec_pts)
            total = family_pts + subfamily_pts + spec_pts
            total = max(75, min(98, total))
            _log(job, cand, family_pts, subfamily_pts, spec_pts, total, "DIRECT")
            return MatchDecision(
                include=True, relation="DIRECT", score=total,
                family_score=family_pts, subfamily_score=subfamily_pts, spec_score=spec_pts,
                job_family=job.family, job_subfamily=job.subfamily,
                candidate_family=cand.family, candidate_subfamily=cand.subfamily,
                reason="Familie si subfamilie identice.",
            )
        else:
            # SAME_FAMILY — subfamilie diferita
            subfamily_pts = max(10, 30 - 10)  # 20/30
            spec_pts = min(10, ai // 10)
            total = family_pts + subfamily_pts + spec_pts
            total = max(55, min(75, total))
            include = not strict or total >= 60
            _log(job, cand, family_pts, subfamily_pts, spec_pts, total, "SAME_FAMILY")
            return MatchDecision(
                include=include, relation="SAME_FAMILY", score=total,
                family_score=family_pts, subfamily_score=subfamily_pts, spec_score=spec_pts,
                job_family=job.family, job_subfamily=job.subfamily,
                candidate_family=cand.family, candidate_subfamily=cand.subfamily,
                reason="Aceeasi familie, subfamilie diferita.",
            )

    # ── Familie adiacenta ────────────────────────────────────────────────────
    if _is_adjacent(job.family, cand.family):
        family_pts = 20   # 20/50 — partial
        subfamily_pts = 0
        spec_pts = 0
        total = max(20, min(40, family_pts + min(10, ai // 10)))
        _log(job, cand, family_pts, subfamily_pts, spec_pts, total, "ADJACENT")
        return MatchDecision(
            include=False if strict else True,
            relation="ADJACENT", score=total,
            family_score=family_pts, subfamily_score=subfamily_pts, spec_score=spec_pts,
            job_family=job.family, job_subfamily=job.subfamily,
            candidate_family=cand.family, candidate_subfamily=cand.subfamily,
            reason="Familie apropiata, dar nu potrivire directa.",
        )

    # ── Familie gresita ──────────────────────────────────────────────────────
    family_pts = 0
    total = min(5, ai // 20)   # maxim 5 puncte, indiferent de AI
    _log(job, cand, family_pts, 0, 0, total, "WRONG_FAMILY")
    return MatchDecision(
        include=False,
        relation="WRONG_FAMILY", score=total,
        family_score=0, subfamily_score=0, spec_score=0,
        job_family=job.family, job_subfamily=job.subfamily,
        candidate_family=cand.family, candidate_subfamily=cand.subfamily,
        reason=f"Familie gresita: job={job.family}, candidat={cand.family}.",
    )


def _log(job: Classification, cand: Classification,
         f: int, s: int, sp: int, total: int, relation: str) -> None:
    logger.debug(
        "MATCH | relation=%s | job_family=%s/%s | cand_family=%s/%s | "
        "family_score=%d | subfamily_score=%d | spec_score=%d | final=%d",
        relation, job.family, job.subfamily,
        cand.family, cand.subfamily,
        f, s, sp, total,
    )


# ─── INTEGRARE CU FLUXUL AI ──────────────────────────────────────────────────

def candidate_text_from_record(candidate) -> str:
    parts = [
        getattr(candidate, "position", ""),
        getattr(candidate, "job_title", ""),
        getattr(candidate, "skills", ""),
        getattr(candidate, "companies", ""),
        getattr(candidate, "experience", ""),
        getattr(candidate, "strengths", ""),
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
    """
    Preia rezultatul AI si il recalibreaza cu motorul deterministic.
    Scorul final respecta formula 50/30/20.
    Familie gresita = maxim 5 puncte.
    """
    role = str(
        data.get("position") or
        data.get("recommended_role_for_candidate") or
        data.get("current_position") or ""
    )
    decision = evaluate_match(
        target_job, cv_text,
        candidate_role=role,
        ai_score=ai_score,
        strict=True,
    )

    data["score"] = decision.score
    data["candidate_family"] = decision.candidate_family
    data["job_family"] = decision.job_family
    data["candidate_subfamily"] = decision.candidate_subfamily
    data["job_subfamily"] = decision.job_subfamily
    data["match_relation"] = decision.relation
    data["family_score"] = decision.family_score
    data["subfamily_score"] = decision.subfamily_score
    data["spec_score"] = decision.spec_score

    if decision.relation == "DIRECT":
        data["recommendation"] = "STRONG_YES" if decision.score >= 85 else "YES"
        data["priority"] = "HIGH" if decision.score >= 85 else "MEDIUM"
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
        f"Potrivire: {decision.relation} | "
        f"Familie: {decision.family_score}/50 | "
        f"Subfamilie: {decision.subfamily_score}/30 | "
        f"Specializare: {decision.spec_score}/20 | "
        f"Job={decision.job_family}/{decision.job_subfamily} | "
        f"CV={decision.candidate_family}/{decision.candidate_subfamily}"
    )
    data["summary"] = (family_line + (" | " + base_summary if base_summary else "")).strip()
    return data
