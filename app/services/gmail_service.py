import os
import imaplib
import email
import re
import time
import unicodedata
from email.header import decode_header

try:
    import resend
except ImportError:
    resend = None

from app.services.cloudinary_service import (
    upload_cv_bytes_to_cloudinary,
    check_cv_exists_on_cloudinary
)

GMAIL_EMAIL = os.getenv("GMAIL_EMAIL")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
GMAIL_IMAP_HOST = os.getenv("GMAIL_IMAP_HOST", "imap.gmail.com")
GMAIL_IMAP_FOLDER = os.getenv("GMAIL_IMAP_FOLDER", "INBOX")

RESEND_API_KEY = os.getenv("RESEND_API_KEY")
FROM_EMAIL = os.getenv("FROM_EMAIL", "onboarding@resend.dev")

ALLOWED_EXTENSIONS = (".pdf", ".docx", ".doc")

# Cuvinte cheie cautate in subiectul mailului
CV_SUBJECT_KEYWORDS = [
    "cv", "c.v", "c.v.", "curriculum vitae", "curriculum",
    "job", "angajare", "angajat", "aplicatie", "aplicatia",
    "candidatura", "candidat", "candidez",
    "dosar", "post", "pozitie", "pozitia",
    "interviu", "recrutare", "resurse umane",
    "trimit cv", "atasez cv", "va trimit",
    "sunt interesat", "ma recomand",
    "referitor la anunt", "in atentia",
    "resume", "application", "hiring", "vacancy", "career",
    "experienta", "competente", "oferta de munca",
    "locul de munca", "loc de munca",
]


def _require_gmail_config():
    if not GMAIL_EMAIL:
        raise RuntimeError("Lipseste GMAIL_EMAIL in Render Environment.")
    if not GMAIL_APP_PASSWORD:
        raise RuntimeError("Lipseste GMAIL_APP_PASSWORD in Render Environment.")


def _decode_mime_value(value):
    if not value:
        return ""
    decoded_parts = decode_header(value)
    result = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            result += part.decode(encoding or "utf-8", errors="ignore")
        else:
            result += part
    return result.strip()


def _safe_filename(filename):
    filename = _decode_mime_value(filename)
    filename = filename.replace("\\", "_").replace("/", "_")
    filename = unicodedata.normalize("NFKD", filename)
    filename = "".join(ch for ch in filename if not unicodedata.combining(ch))
    filename = re.sub(r"[^A-Za-z0-9._() -]+", "_", filename)
    filename = re.sub(r"\s+", " ", filename).strip()
    if not filename:
        filename = f"cv_{int(time.time())}.pdf"
    return filename


def _unique_filename(filename):
    """Genereaza un nume unic cu timestamp pentru a evita coliziuni."""
    base, ext = os.path.splitext(filename)
    timestamp = int(time.time())
    return f"{base}_{timestamp}{ext}"


def subject_contains_cv_keywords(subject):
    """
    Verifica daca subiectul mailului contine cuvinte cheie legate de CV/job.
    Returneaza True daca mailul pare sa fie o candidatura.
    """
    if not subject:
        return False
    subject_lower = subject.lower()
    for keyword in CV_SUBJECT_KEYWORDS:
        if keyword in subject_lower:
            return True
    return False


def filename_is_cv(filename):
    lower_name = (filename or "").lower().strip()
    if not lower_name.endswith(ALLOWED_EXTENSIONS):
        return False
    blocked_keywords = [
        "factura", "invoice", "bon", "chitanta", "chitanta",
        "proforma", "contract", "extras", "plata", "receipt",
        "ordin", "oferta", "gdpr", "acord", "anexa"
    ]
    for blocked in blocked_keywords:
        if blocked in lower_name:
            return False
    return True


def _connect_imap():
    _require_gmail_config()
    mail = imaplib.IMAP4_SSL(GMAIL_IMAP_HOST, 993)
    mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
    mail.select(GMAIL_IMAP_FOLDER)
    return mail


def _iter_message_parts(message):
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_maintype() == "multipart":
                continue
            yield part
    else:
        yield message


def _process_message(mail, message_id, downloaded, skipped):
    """
    Proceseaza un singur mail:
    1. Verifica subiectul — daca nu contine cuvinte cheie CV, skip
    2. Extrage atasamentele CV
    3. Verifica duplicatele pe Cloudinary
    4. Uploadeaza direct pe Cloudinary din bytes (fara disk local)
    """
    status, data = mail.fetch(message_id, "(RFC822)")
    if status != "OK" or not data or not data[0]:
        return

    raw_email = data[0][1]
    message = email.message_from_bytes(raw_email)

    # Filtreaza dupa subiect
    subject = _decode_mime_value(message.get("Subject", ""))
    if not subject_contains_cv_keywords(subject):
        return  # Mail irelevant, skip complet

    for part in _iter_message_parts(message):
        disposition = part.get_content_disposition()
        filename = part.get_filename()

        if disposition != "attachment" and not filename:
            continue
        if not filename:
            continue

        filename = _safe_filename(filename)

        if not filename_is_cv(filename):
            skipped.append(filename)
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            skipped.append(filename)
            continue

        # Verifica duplicat pe Cloudinary
        if check_cv_exists_on_cloudinary(filename):
            skipped.append(filename)
            continue

        # Daca exista un fisier cu acelasi nume, adauga timestamp
        if check_cv_exists_on_cloudinary(filename):
            filename = _unique_filename(filename)

        # Upload direct pe Cloudinary din bytes — fara disk local
        try:
            upload_cv_bytes_to_cloudinary(payload, filename)
            downloaded.append(filename)
        except Exception as e:
            skipped.append(f"{filename} (eroare: {str(e)[:50]})")


def download_cv_attachments(max_results=50):
    """
    Cauta in Gmail mailuri cu subiect relevant (CV, job, angajare etc.)
    si uploadeaza atasamentele noi direct pe Cloudinary.

    - Nu salveaza nimic pe disk local
    - Verifica duplicatele pe Cloudinary
    - Filtreaza dupa subiectul mailului
    """
    mail = None
    downloaded = []
    skipped = []

    try:
        mail = _connect_imap()

        # Cauta mailuri cu atasamente din ultimele 60 zile
        try:
            status, data = mail.search(
                None,
                "X-GM-RAW",
                '"has:attachment newer_than:60d"'
            )
        except imaplib.IMAP4.error:
            status, data = mail.search(None, "ALL")

        if status != "OK":
            status, data = mail.search(None, "ALL")

        ids = data[0].split() if data and data[0] else []
        ids = ids[-max_results:]
        ids.reverse()

        for message_id in ids:
            _process_message(mail, message_id, downloaded, skipped)

    finally:
        if mail:
            try:
                mail.close()
            except Exception:
                pass
            try:
                mail.logout()
            except Exception:
                pass

    return {
        "downloaded": downloaded,
        "skipped": skipped,
        "downloaded_count": len(downloaded),
        "skipped_count": len(skipped)
    }


def send_email_api(to_email: str, subject: str, body: str):
    if resend is None:
        raise RuntimeError("Lipseste pachetul resend.")
    if not RESEND_API_KEY:
        raise RuntimeError("Lipseste RESEND_API_KEY in Render Environment.")
    if not FROM_EMAIL:
        raise RuntimeError("Lipseste FROM_EMAIL in Render Environment.")
    if not to_email:
        raise RuntimeError("Lipseste adresa destinatarului.")
    resend.api_key = RESEND_API_KEY
    return resend.Emails.send({
        "from": FROM_EMAIL,
        "to": [to_email],
        "subject": subject,
        "text": body
    })


def send_interview_email(to_email, candidate_name, job_title):
    if not to_email:
        raise ValueError("Candidatul nu are email extras din CV.")
    safe_name = candidate_name or "candidat"
    safe_job = job_title or "postul pentru care ati aplicat"
    subject = f"Invitatie la interviu - {safe_job}"
    body = f"""Buna ziua, {safe_name},

Va multumim pentru CV-ul transmis pentru postul de {safe_job}.

In urma analizei initiale, dorim sa continuam procesul de recrutare.

Va rugam sa ne transmiteti disponibilitatea dumneavoastra pentru o prima discutie.

Cu respect,
Echipa HR
"""
    result = send_email_api(to_email=to_email, subject=subject, body=body)
    return {"sent": True, "provider": "resend", "result": result, "to": to_email, "subject": subject}


def send_rejection_email(to_email, candidate_name, job_title):
    if not to_email:
        raise ValueError("Candidatul nu are email extras din CV.")
    safe_name = candidate_name or "candidat"
    safe_job = job_title or "postul pentru care ati aplicat"
    subject = f"Raspuns privind procesul de recrutare - {safe_job}"
    body = f"""Buna ziua, {safe_name},

Va multumim pentru CV-ul transmis si pentru interesul acordat postului de {safe_job}.

In urma analizei initiale, am decis sa continuam procesul de recrutare cu alti candidati ale caror experiente sunt mai apropiate de cerintele actuale ale rolului.

Vom pastra datele dumneavoastra pentru oportunitati viitoare potrivite profilului transmis.

Va multumim pentru timpul acordat si va dorim mult succes in continuare.

Cu respect,
Echipa HR
"""
    result = send_email_api(to_email=to_email, subject=subject, body=body)
    return {"sent": True, "provider": "resend", "result": result, "to": to_email, "subject": subject}


if __name__ == "__main__":
    result = download_cv_attachments()
    print("CV-uri descarcate:", result["downloaded_count"])
    for f in result["downloaded"]:
        print("  +", f)
    print("Sarite:", result["skipped_count"])
