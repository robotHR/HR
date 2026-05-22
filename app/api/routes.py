import os
import shutil
import time
from collections import Counter, defaultdict
from io import BytesIO

from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, FileResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy import inspect, text
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from app.core.database import SessionLocal, engine
from app.models.candidate_model import Candidate
from app.models.candidate_event_model import CandidateEvent
from app.services.cv_parser import process_cvs_for_job
from app.services.gmail_service import (
    download_cv_attachments,
    send_interview_email,
    send_rejection_email
)
from app.services.multi_job_analyzer import MultiJobAnalyzer

router = APIRouter()
templates = Jinja2Templates(directory="app/templates")

UPLOAD_FOLDER = "app/uploads"
TERMINAL_STATUSES = ["REFUZAT", "ANGAJAT", "EXCLUS"]


def ensure_candidate_extra_columns():
    inspector = inspect(engine)

    if not inspector.has_table("candidates"):
        return

    existing = [column["name"] for column in inspector.get_columns("candidates")]

    with engine.begin() as conn:
        if "batch_id" not in existing:
            conn.execute(text("ALTER TABLE candidates ADD COLUMN batch_id VARCHAR"))

        if "visible_in_dashboard" not in existing:
            conn.execute(text("ALTER TABLE candidates ADD COLUMN visible_in_dashboard INTEGER DEFAULT 1"))
            conn.execute(text("UPDATE candidates SET visible_in_dashboard = 1 WHERE visible_in_dashboard IS NULL"))


ensure_candidate_extra_columns()


def load_candidates():
    db = SessionLocal()
    candidates = db.query(Candidate).order_by(Candidate.id.desc()).all()
    total = len(candidates)
    db.close()
    return candidates, total


def load_dashboard_candidates():
    db = SessionLocal()
    candidates = db.query(Candidate).filter(
        Candidate.visible_in_dashboard == 1
    ).filter(
        Candidate.status.notin_(TERMINAL_STATUSES)
    ).order_by(
        Candidate.id.desc()
    ).all()
    total = len(candidates)
    db.close()
    return candidates, total


def get_candidate_by_id(candidate_id):
    db = SessionLocal()
    candidate = db.query(Candidate).filter(Candidate.id == candidate_id).first()
    db.close()
    return candidate


def get_candidate_events(candidate_id):
    db = SessionLocal()
    events = db.query(CandidateEvent).filter(
        CandidateEvent.candidate_id == candidate_id
    ).order_by(CandidateEvent.id.desc()).all()
    db.close()
    return events


def get_last_candidate_event(candidate_id):
    db = SessionLocal()
    event = db.query(CandidateEvent).filter(
        CandidateEvent.candidate_id == candidate_id
    ).order_by(CandidateEvent.id.desc()).first()
    db.close()
    return event


def log_candidate_event(candidate_id, event_type, title, description=""):
    db = SessionLocal()
    event = CandidateEvent(
        candidate_id=candidate_id,
        event_type=event_type,
        title=title,
        description=description
    )
    db.add(event)
    db.commit()
    db.close()


def get_max_candidate_id():
    db = SessionLocal()
    candidate = db.query(Candidate).order_by(Candidate.id.desc()).first()
    max_id = candidate.id if candidate else 0
    db.close()
    return max_id


def get_final_cv_files():
    db = SessionLocal()
    rows = db.query(Candidate).filter(
        Candidate.status.in_(TERMINAL_STATUSES)
    ).all()
    files = {row.cv_file for row in rows if row.cv_file}
    db.close()
    return files


def mark_new_candidates_batch(after_id, batch_id, visible=1, only_files=None, hide_final_files=True):
    final_files = get_final_cv_files() if hide_final_files else set()
    db = SessionLocal()
    rows = db.query(Candidate).filter(Candidate.id > after_id).all()

    for row in rows:
        row.batch_id = batch_id

        should_show = bool(visible)

        if only_files is not None and row.cv_file not in only_files:
            should_show = False

        if row.cv_file in final_files:
            should_show = False

        if row.status in TERMINAL_STATUSES:
            should_show = False

        row.visible_in_dashboard = 1 if should_show else 0

    db.commit()
    db.close()


def clear_dashboard_only():
    db = SessionLocal()
    rows = db.query(Candidate).all()

    for row in rows:
        row.visible_in_dashboard = 0

    db.commit()
    db.close()


def update_candidate_status(candidate_id, status, event_title=None, event_description=""):
    db = SessionLocal()
    candidate = db.query(Candidate).filter(Candidate.id == candidate_id).first()

    if not candidate:
        db.close()
        return None

    old_status = candidate.status or "NOU"
    candidate.status = status

    if status in TERMINAL_STATUSES:
        candidate.visible_in_dashboard = 0

    db.commit()
    db.refresh(candidate)
    db.close()

    log_candidate_event(
        candidate_id=candidate_id,
        event_type=status,
        title=event_title or f"Status schimbat: {status_display(old_status)} -> {status_display(status)}",
        description=event_description
    )

    return candidate


def update_candidate_manual(candidate_id, name, email, phone, status, position, experience, skills, summary):
    db = SessionLocal()
    candidate = db.query(Candidate).filter(Candidate.id == candidate_id).first()

    if not candidate:
        db.close()
        return None

    old_status = candidate.status or "NOU"

    candidate.name = name.strip() if name else candidate.name
    candidate.email = email.strip() if email else ""
    candidate.phone = phone.strip() if phone else ""
    candidate.status = status.strip() if status else candidate.status
    candidate.position = position.strip() if position else ""
    candidate.experience = experience.strip() if experience else ""
    candidate.skills = skills.strip() if skills else ""
    candidate.summary = summary.strip() if summary else ""

    if candidate.status in TERMINAL_STATUSES:
        candidate.visible_in_dashboard = 0

    new_status = candidate.status or "NOU"

    db.commit()
    db.refresh(candidate)
    db.close()

    log_candidate_event(
        candidate_id=candidate_id,
        event_type="EDIT",
        title="Date candidat actualizate manual",
        description=f"Status anterior: {status_display(old_status)}. Status actual: {status_display(new_status)}."
    )

    return candidate


def group_candidates_by_job(candidates):
    jobs = {}

    for candidate in candidates:
        job_name = candidate.job_title or "Fara categorie"

        if job_name not in jobs:
            jobs[job_name] = {"latest_id": candidate.id, "candidates": []}

        jobs[job_name]["latest_id"] = max(jobs[job_name]["latest_id"], candidate.id)
        jobs[job_name]["candidates"].append(candidate)

    sorted_jobs = sorted(jobs.items(), key=lambda item: item[1]["latest_id"], reverse=True)

    final_jobs = {}
    for job_name, data in sorted_jobs:
        final_jobs[job_name] = sorted(data["candidates"], key=lambda candidate: candidate.score or 0, reverse=True)

    return final_jobs


def extract_risk(summary, risk_name):
    text_value = (summary or "").lower()
    marker = risk_name.lower() + ":"

    if marker not in text_value:
        return "unknown"

    after = text_value.split(marker, 1)[1].strip()

    if after.startswith("low"):
        return "low"
    if after.startswith("medium"):
        return "medium"
    if after.startswith("high"):
        return "high"

    return "unknown"


def risk_label(value):
    value = (value or "unknown").lower()
    if value == "low":
        return "scazut"
    if value == "medium":
        return "mediu"
    if value == "high":
        return "ridicat"
    return "necunoscut"


def status_display(status):
    status = status or "NOU"

    mapping = {
        "ADMIS": "Top Candidat",
        "DE ANALIZAT": "Potential Candidat",
        "RESPINS": "Potrivit altor roluri",
        "SELECTAT": "Potential Candidat",
        "INTERVIU": "Interviu",
        "REFUZAT": "Potrivit altor roluri",
        "ANGAJAT": "Angajat",
        "EXCLUS": "Candidat Exclus",
        "NOU": "Nou"
    }

    return mapping.get(status, status)


def decision_text(candidate):
    score = candidate.score or 0
    status = candidate.status or "NOU"

    messages = {
        "INTERVIU": "Email trimis. Candidatul trebuie contactat pentru prima discutie.",
        "REFUZAT": "Candidatul a fost inchis pentru acest rol si poate ramane potrivit pentru alte roluri.",
        "EXCLUS": "Candidat exclus din proces. Nu se trimite email si nu mai apare in dashboard.",
        "ANGAJAT": "Candidatul a fost marcat ca angajat.",
        "SELECTAT": "Candidatul are potential si a fost selectat pentru urmatorul pas.",
        "ADMIS": "Top candidat pentru acest rol. Verifica detaliile si disponibilitatea.",
        "DE ANALIZAT": "Potential candidat. Necesita validare telefonica.",
        "RESPINS": "Candidatul pare mai potrivit pentru alte roluri."
    }

    if status in messages:
        return messages[status]

    if score >= 70:
        return "Profil cu potrivire buna."
    if score >= 45:
        return "Profil intermediar. Necesita verificare."
    return "Profil slab potrivit pentru acest rol."


def status_class(status):
    status = status or "NOU"

    if status in ["ADMIS", "ANGAJAT"]:
        return "green"
    if status in ["DE ANALIZAT", "SELECTAT", "INTERVIU"]:
        return "yellow"
    if status in ["RESPINS", "REFUZAT", "EXCLUS"]:
        return "red"
    return "default"


def event_class(event_type):
    event_type = event_type or "INFO"

    if event_type in ["EMAIL_INTERVIU", "ANGAJAT", "INTERVIU"]:
        return "green"
    if event_type in ["EMAIL_REFUZ", "REFUZAT", "EXCLUS"]:
        return "red"
    if event_type in ["STATUS", "EDIT", "SELECTAT", "NOTE", "DE ANALIZAT"]:
        return "yellow"
    return "default"


def format_event_date(value):
    if not value:
        return ""
    try:
        return value.strftime("%d.%m.%Y %H:%M")
    except Exception:
        return str(value)


def calc_percent(part, total):
    if not total:
        return 0
    return round((part / total) * 100, 1)


def temporarily_keep_only_files(files_to_keep):
    hold_folder = os.path.join(UPLOAD_FOLDER, "_hidden_temp")
    os.makedirs(hold_folder, exist_ok=True)
    moved = []

    for filename in os.listdir(UPLOAD_FOLDER):
        full_path = os.path.join(UPLOAD_FOLDER, filename)

        if not os.path.isfile(full_path):
            continue

        if filename not in files_to_keep:
            target = os.path.join(hold_folder, filename)
            shutil.move(full_path, target)
            moved.append((target, full_path))

    return moved


def restore_temp_files(moved_files):
    for old_path, new_path in moved_files:
        if os.path.exists(old_path):
            shutil.move(old_path, new_path)


def build_analytics(candidates):
    total = len(candidates)
    status_counter = Counter(candidate.status or "NOU" for candidate in candidates)
    scores = [candidate.score or 0 for candidate in candidates]
    avg_score = round(sum(scores) / total, 1) if total else 0

    high_score = len([score for score in scores if score >= 70])
    medium_score = len([score for score in scores if 45 <= score < 70])
    low_score = len([score for score in scores if score < 45])

    job_stats = []
    grouped = defaultdict(list)

    for candidate in candidates:
        grouped[candidate.job_title or "Fara categorie"].append(candidate)

    for job_title, items in grouped.items():
        job_scores = [item.score or 0 for item in items]
        job_status = Counter(item.status or "NOU" for item in items)

        job_stats.append({
            "job_title": job_title,
            "total": len(items),
            "avg_score": round(sum(job_scores) / len(items), 1) if items else 0,
            "admis": job_status.get("ADMIS", 0),
            "de_analizat": job_status.get("DE ANALIZAT", 0),
            "interviu": job_status.get("INTERVIU", 0),
            "selectat": job_status.get("SELECTAT", 0),
            "refuzat": job_status.get("REFUZAT", 0),
            "exclus": job_status.get("EXCLUS", 0),
            "angajat": job_status.get("ANGAJAT", 0),
            "respins": job_status.get("RESPINS", 0),
            "top_candidate": max(items, key=lambda item: item.score or 0) if items else None
        })

    job_stats = sorted(job_stats, key=lambda item: item["total"], reverse=True)
    top_candidates = sorted(candidates, key=lambda candidate: candidate.score or 0, reverse=True)[:10]

    pipeline = [
        {"label": "Nou", "count": status_counter.get("NOU", 0)},
        {"label": "Potential Candidat", "count": status_counter.get("DE ANALIZAT", 0) + status_counter.get("SELECTAT", 0)},
        {"label": "Interviu", "count": status_counter.get("INTERVIU", 0)},
        {"label": "Top Candidat", "count": status_counter.get("ADMIS", 0)},
        {"label": "Angajat", "count": status_counter.get("ANGAJAT", 0)},
        {"label": "Potrivit altor roluri", "count": status_counter.get("REFUZAT", 0) + status_counter.get("RESPINS", 0)},
        {"label": "Candidat Exclus", "count": status_counter.get("EXCLUS", 0)}
    ]

    return {
        "total": total,
        "avg_score": avg_score,
        "high_score": high_score,
        "medium_score": medium_score,
        "low_score": low_score,
        "status_counter": status_counter,
        "job_stats": job_stats,
        "top_candidates": top_candidates,
        "pipeline": pipeline,
        "rates": {
            "interview_rate": calc_percent(status_counter.get("INTERVIU", 0), total),
            "hire_rate": calc_percent(status_counter.get("ANGAJAT", 0), total),
            "reject_rate": calc_percent(
                status_counter.get("REFUZAT", 0) + status_counter.get("RESPINS", 0) + status_counter.get("EXCLUS", 0),
                total
            ),
            "shortlist_rate": calc_percent(
                status_counter.get("ADMIS", 0)
                + status_counter.get("DE ANALIZAT", 0)
                + status_counter.get("SELECTAT", 0)
                + status_counter.get("INTERVIU", 0)
                + status_counter.get("ANGAJAT", 0),
                total
            )
        }
    }


def apply_candidate_filters(candidates, q="", status="", job="", min_score=0):
    filtered = candidates

    if q:
        search = q.lower().strip()
        filtered = [
            candidate for candidate in filtered
            if search in (candidate.name or "").lower()
            or search in (candidate.email or "").lower()
            or search in (candidate.phone or "").lower()
            or search in (candidate.cv_file or "").lower()
        ]

    if status:
        filtered = [candidate for candidate in filtered if (candidate.status or "") == status]

    if job:
        filtered = [candidate for candidate in filtered if (candidate.job_title or "") == job]

    if min_score:
        filtered = [candidate for candidate in filtered if (candidate.score or 0) >= min_score]

    return filtered


def candidate_last_event_text(candidate_id):
    event = get_last_candidate_event(candidate_id)

    if not event:
        return "", ""

    return event.title or "", format_event_date(event.created_at)


def generate_candidates_excel(candidates):
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidati"

    headers = [
        "Nume",
        "Email",
        "Telefon",
        "Job cautat",
        "Pozitie detectata",
        "Scor",
        "Status",
        "Experienta",
        "Skill-uri",
        "CV fisier",
        "Rezumat AI",
        "Ultima actiune",
        "Data ultimei actiuni"
    ]

    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for candidate in candidates:
        last_action, last_action_date = candidate_last_event_text(candidate.id)

        ws.append([
            candidate.name or "",
            candidate.email or "",
            candidate.phone or "",
            candidate.job_title or "",
            candidate.position or "",
            candidate.score or 0,
            status_display(candidate.status),
            candidate.experience or "",
            candidate.skills or "",
            candidate.cv_file or "",
            candidate.summary or "",
            last_action,
            last_action_date
        ])

    widths = {
        "A": 28, "B": 34, "C": 18, "D": 28, "E": 28, "F": 10,
        "G": 24, "H": 20, "I": 45, "J": 32, "K": 65, "L": 28, "M": 22
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output


templates.env.globals["extract_risk"] = extract_risk
templates.env.globals["risk_label"] = risk_label
templates.env.globals["decision_text"] = decision_text
templates.env.globals["status_class"] = status_class
templates.env.globals["status_display"] = status_display
templates.env.globals["event_class"] = event_class
templates.env.globals["format_event_date"] = format_event_date
templates.env.globals["calc_percent"] = calc_percent


@router.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    candidates, total = load_dashboard_candidates()
    jobs = group_candidates_by_job(candidates)
    message = request.query_params.get("message")

    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={"request": request, "candidates": candidates, "jobs": jobs, "total": total, "message": message}
    )


@router.get("/candidati", response_class=HTMLResponse)
async def candidates_page(request: Request, q: str = "", status: str = "", job: str = "", min_score: int = 0):
    candidates, total = load_candidates()
    all_jobs = sorted(list({candidate.job_title for candidate in candidates if candidate.job_title}))
    all_statuses = sorted(list({candidate.status for candidate in candidates if candidate.status}))

    filtered = apply_candidate_filters(candidates, q=q, status=status, job=job, min_score=min_score)

    stats = {
        "total": len(candidates),
        "filtered": len(filtered),
        "interviu": len([c for c in candidates if c.status == "INTERVIU"]),
        "selectat": len([c for c in candidates if c.status == "SELECTAT"]),
        "angajat": len([c for c in candidates if c.status == "ANGAJAT"]),
        "refuzat": len([c for c in candidates if c.status == "REFUZAT"]),
        "exclus": len([c for c in candidates if c.status == "EXCLUS"]),
        "de_analizat": len([c for c in candidates if c.status == "DE ANALIZAT"]),
    }

    return templates.TemplateResponse(
        request=request,
        name="candidates.html",
        context={
            "request": request,
            "candidates": filtered,
            "total": total,
            "stats": stats,
            "all_jobs": all_jobs,
            "all_statuses": all_statuses,
            "q": q,
            "status": status,
            "job": job,
            "min_score": min_score
        }
    )


@router.get("/candidati/export")
async def export_candidates(q: str = "", status: str = "", job: str = "", min_score: int = 0):
    candidates, total = load_candidates()
    filtered = apply_candidate_filters(candidates, q=q, status=status, job=job, min_score=min_score)

    output = generate_candidates_excel(filtered)
    filename = f"nexas_hr_candidati_{int(time.time())}.xlsx"

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename={filename}"
        }
    )


@router.get("/analytics", response_class=HTMLResponse)
async def analytics_page(request: Request):
    candidates, total = load_candidates()
    analytics = build_analytics(candidates)

    return templates.TemplateResponse(
        request=request,
        name="analytics.html",
        context={"request": request, "candidates": candidates, "total": total, "analytics": analytics}
    )


@router.get("/candidat/{candidate_id}", response_class=HTMLResponse)
async def candidate_detail(request: Request, candidate_id: int):
    candidate = get_candidate_by_id(candidate_id)

    if not candidate:
        return templates.TemplateResponse(
            request=request,
            name="candidate_detail.html",
            context={"request": request, "candidate": None, "events": [], "message": "Candidatul nu a fost gasit."},
            status_code=404
        )

    message = request.query_params.get("message")
    error = request.query_params.get("error")
    events = get_candidate_events(candidate_id)

    return templates.TemplateResponse(
        request=request,
        name="candidate_detail.html",
        context={"request": request, "candidate": candidate, "events": events, "message": message, "error": error}
    )


@router.post("/candidat/{candidate_id}/salveaza")
async def save_candidate_manual(
    candidate_id: int,
    name: str = Form(""),
    email: str = Form(""),
    phone: str = Form(""),
    status: str = Form(""),
    position: str = Form(""),
    experience: str = Form(""),
    skills: str = Form(""),
    summary: str = Form("")
):
    candidate = update_candidate_manual(candidate_id, name, email, phone, status, position, experience, skills, summary)

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    return RedirectResponse(
        url=f"/candidat/{candidate_id}?message=Datele candidatului au fost actualizate.",
        status_code=303
    )


@router.post("/candidat/{candidate_id}/nota")
async def add_candidate_note(candidate_id: int, note: str = Form("")):
    candidate = get_candidate_by_id(candidate_id)

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    note = note.strip()

    if not note:
        return RedirectResponse(
            url=f"/candidat/{candidate_id}?error=Nota HR este goala.",
            status_code=303
        )

    log_candidate_event(
        candidate_id=candidate_id,
        event_type="NOTE",
        title="Nota HR adaugata",
        description=note
    )

    return RedirectResponse(
        url=f"/candidat/{candidate_id}?message=Nota HR a fost salvata.",
        status_code=303
    )


@router.get("/cv/{candidate_id}")
async def open_candidate_cv(candidate_id: int):
    candidate = get_candidate_by_id(candidate_id)

    if not candidate or not candidate.cv_file:
        return JSONResponse({"error": "CV-ul nu a fost gasit pentru acest candidat."}, status_code=404)

    file_path = os.path.join("app", "uploads", candidate.cv_file)

    if not os.path.exists(file_path):
        return JSONResponse({"error": "Fisierul CV nu exista in app/uploads."}, status_code=404)

    return FileResponse(path=file_path, filename=candidate.cv_file, media_type="application/pdf")


@router.post("/candidat/{candidate_id}/trimite-interviu")
async def send_interview(candidate_id: int):
    candidate = get_candidate_by_id(candidate_id)

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    if not candidate.email:
        return RedirectResponse(url=f"/candidat/{candidate_id}?error=Candidatul nu are email extras din CV.", status_code=303)

    try:
        send_interview_email(to_email=candidate.email, candidate_name=candidate.name, job_title=candidate.job_title)
        update_candidate_status(
            candidate_id,
            "INTERVIU",
            event_title="Email interviu trimis",
            event_description=f"Email trimis catre {candidate.email}."
        )

        return RedirectResponse(
            url=f"/candidat/{candidate_id}?message=Email trimis. Status actualizat: Interviu.",
            status_code=303
        )

    except Exception as e:
        return RedirectResponse(
            url=f"/candidat/{candidate_id}?error=Eroare la trimiterea emailului: {str(e)}",
            status_code=303
        )


@router.post("/candidat/{candidate_id}/refuza")
async def reject_candidate(candidate_id: int):
    candidate = get_candidate_by_id(candidate_id)

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    if not candidate.email:
        update_candidate_status(
            candidate_id,
            "REFUZAT",
            event_title="Candidat mutat la Potrivit altor roluri fara email",
            event_description="Candidatul nu are email extras din CV. Nu s-a trimis email."
        )
        return RedirectResponse(
            url=f"/candidat/{candidate_id}?message=Status actualizat: Potrivit altor roluri. Candidatul nu are email extras din CV, deci nu s-a trimis email.",
            status_code=303
        )

    try:
        send_rejection_email(to_email=candidate.email, candidate_name=candidate.name, job_title=candidate.job_title)
        update_candidate_status(
            candidate_id,
            "REFUZAT",
            event_title="Email de refuz trimis",
            event_description=f"Email trimis catre {candidate.email}."
        )

        return RedirectResponse(
            url=f"/candidat/{candidate_id}?message=Email de refuz trimis. Status actualizat: Potrivit altor roluri.",
            status_code=303
        )

    except Exception as e:
        return RedirectResponse(
            url=f"/candidat/{candidate_id}?error=Eroare la trimiterea emailului de refuz: {str(e)}",
            status_code=303
        )


@router.post("/candidat/{candidate_id}/exclude")
async def exclude_candidate(candidate_id: int, reason: str = Form("")):
    candidate = get_candidate_by_id(candidate_id)

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    reason = reason.strip() or "Candidat exclus manual din proces. Nu s-a trimis email."

    update_candidate_status(
        candidate_id,
        "EXCLUS",
        event_title="Candidat exclus",
        event_description=reason
    )

    return RedirectResponse(
        url=f"/candidat/{candidate_id}?message=Candidatul a fost exclus. Nu s-a trimis email.",
        status_code=303
    )


@router.post("/candidat/{candidate_id}/angajeaza")
async def hire_candidate(candidate_id: int):
    candidate = update_candidate_status(
        candidate_id,
        "ANGAJAT",
        event_title="Candidat marcat ca angajat",
        event_description="Decizie finala pozitiva."
    )

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    return RedirectResponse(url=f"/candidat/{candidate_id}?message=Status actualizat: Angajat.", status_code=303)


@router.post("/candidat/{candidate_id}/selecteaza")
async def select_candidate(candidate_id: int):
    candidate = update_candidate_status(
        candidate_id,
        "SELECTAT",
        event_title="Potential candidat selectat",
        event_description="Candidatul a fost selectat manual pentru urmatorul pas."
    )

    if not candidate:
        return RedirectResponse(url="/?message=Candidatul nu a fost gasit.", status_code=303)

    return RedirectResponse(url=f"/candidat/{candidate_id}?message=Status actualizat: Potential Candidat.", status_code=303)


@router.post("/analizeaza-job")
async def analyze_job(target_job: str = Form(...)):
    batch_id = str(int(time.time()))
    after_id = get_max_candidate_id()
    final_files = get_final_cv_files()

    result = process_cvs_for_job(target_job)

    mark_new_candidates_batch(
        after_id=after_id,
        batch_id=batch_id,
        visible=1,
        only_files=None,
        hide_final_files=True
    )

    message = result.get("message", "Analiza finalizata.")
    if final_files:
        message += " Candidatii marcati Angajat, Exclus sau Potrivit altor roluri nu mai apar in dashboard."

    return RedirectResponse(url=f"/?message={message}", status_code=303)


@router.post("/gmail-si-analiza")
async def gmail_and_analyze(target_job: str = Form(...)):
    batch_id = str(int(time.time()))
    after_id = get_max_candidate_id()

    gmail_result = download_cv_attachments(max_results=30)
    downloaded_files = set(gmail_result.get("downloaded", []))

    if not downloaded_files:
        return RedirectResponse(
            url="/?message=Gmail verificat. Nu exista CV-uri noi descarcate.",
            status_code=303
        )

    moved_files = temporarily_keep_only_files(downloaded_files)

    try:
        analysis_result = process_cvs_for_job(target_job)
    finally:
        restore_temp_files(moved_files)

    mark_new_candidates_batch(
        after_id=after_id,
        batch_id=batch_id,
        visible=1,
        only_files=downloaded_files,
        hide_final_files=True
    )

    message = (
        f"Gmail verificat. "
        f"CV-uri noi descarcate: {gmail_result['downloaded_count']}. "
        f"CV-uri deja existente: {gmail_result['skipped_count']}. "
        f"{analysis_result.get('message', '')}"
    )

    return RedirectResponse(url=f"/?message={message}", status_code=303)


@router.post("/sterge-baza")
async def clear_database():
    clear_dashboard_only()
    return RedirectResponse(
        url="/?message=Dashboard curatat. Candidatii raman salvati in pagina Candidati.",
        status_code=303
    )


@router.post("/sterge-definitiv")
async def delete_all_database():
    db = SessionLocal()
    db.query(CandidateEvent).delete()
    db.query(Candidate).delete()
    db.commit()
    db.close()

    return RedirectResponse(url="/?message=Toata baza de date a fost stearsa definitiv.", status_code=303)


@router.get("/api/debug")
async def debug_db():
    candidates, total = load_candidates()

    return JSONResponse({
        "total_candidates": total,
        "items": [
            {
                "id": candidate.id,
                "name": candidate.name,
                "score": candidate.score,
                "status": candidate.status,
                "status_display": status_display(candidate.status),
                "visible_in_dashboard": candidate.visible_in_dashboard,
                "batch_id": candidate.batch_id,
                "job_title": candidate.job_title,
                "cv_file": candidate.cv_file
            }
            for candidate in candidates
        ]
    })

@router.post("/analyze-multi-jobs")
async def analyze_multi_jobs(request: Request, jobs_input: str = Form(...)):
    """
    Analiză multi-job
    Input: "Casier, Sofer TIR, Lucrator Depozit"
    Output: Ranking-uri per job cu 7-score breakdown
    """
    try:
        analyzer = MultiJobAnalyzer()
        job_list = [j.strip() for j in jobs_input.split(",")]
        results = analyzer.analyze_all_cvs_for_jobs(job_list)
        
        # Export la CSV
        csv_path = analyzer.export_to_csv(results, f"app/static/analysis_results.csv")
        
        return templates.TemplateResponse("analysis_results.html", {
            "request": request,
            "results": results,
            "csv_path": csv_path
        })
    except Exception as e:
        return templates.TemplateResponse("error.html", {
            "request": request,
            "error": str(e)
        })
