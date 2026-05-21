import os
import re
import imaplib
import smtplib
from email import message_from_bytes
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import Dict, List


UPLOAD_FOLDER = "app/uploads"

GMAIL_EMAIL = os.getenv("GMAIL_EMAIL")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

IMAP_HOST = os.getenv("GMAIL_IMAP_HOST", "imap.gmail.com")
IMAP_FOLDER = os.getenv("GMAIL_IMAP_FOLDER", "INBOX")
SMTP_HOST = os.getenv("GMAIL_SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("GMAIL_SMTP_PORT", "465"))

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def _get_gmail_credentials():
    email_address = (os.getenv("GMAIL_EMAIL") or "").strip()
    app_password = (os.getenv("GMAIL_APP_PASSWORD") or "").replace(" ", "").strip()

    if not email_address:
        raise RuntimeError("Lipseste GMAIL_EMAIL in Render Environment.")

    if not app_password:
        raise RuntimeError("Lipseste GMAIL_APP_PASSWORD in Render Environment.")

    return email_address, app_password


def _decode_mime_value(value: str) -> str:
    if not value:
        return ""

    decoded_parts = decode_header(value)
    final_value = ""

    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            final_value += part.decode(encoding or "utf-8", errors="ignore")
        else:
            final_value += part

    return final_value.strip()


def _safe_filename(filename: str) -> str:
    filename = _decode_mime_value(filename)
    filename = os.path.basename(filename)
    filename = filename.replace("\\", "_").replace("/", "_")
    filename = re.sub(r"[^A-Za-z0-9ăâîșțĂÂÎȘȚ._ -]", "_", filename)
    filename = re.sub(r"\s+", " ", filename).strip()

    if not filename:
        filename = "cv_fara_nume.pdf"

    return filename


def _unique_filepath(folder: str, filename: str):
    base, ext = os.path.splitext(filename)
    path = os.path.join(folder, filename)

    if not os.path.exists(path):
        return path, filename, False

    counter = 2
    while True:
        new_filename = f"{base}_{counter}{ext}"
        new_path = os.path.join(folder, new_filename)

        if not os.path.exists(new_path):
            return new_path, new_filename, False

        counter += 1


def filename_is_cv(filename: str) -> bool:
    lower_name = (filename or "").lower().strip()

    if not lower_name.endswith((".pdf", ".docx", ".doc")):
        return False

    blocked_keywords = [
        "factura",
        "invoice",
        "bon",
        "chitanta",
        "chitanță",
        "proforma",
        "contract",
        "extras",
        "plata",
        "receipt",
        "oferta",
        "comanda"
    ]

    return not any(blocked in lower_name for blocked in blocked_keywords)


def _connect_imap():
    email_address, app_password = _get_gmail_credentials()

    mail = imaplib.IMAP4_SSL(IMAP_HOST)
    mail.login(email_address, app_password)

    status, _ = mail.select(IMAP_FOLDER)

    if status != "OK":
        mail.logout()
        raise RuntimeError(f"Nu pot deschide folderul Gmail: {IMAP_FOLDER}")

    return mail


def download_cv_attachments(max_results: int = 20) -> Dict[str, object]:
    """
    Citeste Gmail prin IMAP si descarca atasamente PDF/DOCX/DOC in app/uploads.

    Returneaza acelasi format folosit de routes.py:
    {
        "downloaded": [...],
        "skipped": [...],
        "downloaded_count": int,
        "skipped_count": int
    }
    """
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    downloaded: List[str] = []
    skipped: List[str] = []

    mail = _connect_imap()

    try:
        status, data = mail.search(None, "ALL")

        if status != "OK":
            raise RuntimeError("Nu pot cauta emailuri in Gmail prin IMAP.")

        message_ids = data[0].split()

        if not message_ids:
            return {
                "downloaded": downloaded,
                "skipped": skipped,
                "downloaded_count": 0,
                "skipped_count": 0
            }

        latest_ids = list(reversed(message_ids))[:max_results]

        for message_id in latest_ids:
            status, msg_data = mail.fetch(message_id, "(RFC822)")

            if status != "OK" or not msg_data:
                continue

            raw_email = None
            for item in msg_data:
                if isinstance(item, tuple):
                    raw_email = item[1]
                    break

            if not raw_email:
                continue

            message = message_from_bytes(raw_email)

            for part in message.walk():
                content_disposition = str(part.get("Content-Disposition") or "").lower()
                filename = part.get_filename()

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

                filepath, saved_filename, already_exists = _unique_filepath(UPLOAD_FOLDER, filename)

                if already_exists:
                    skipped.append(filename)
                    continue

                with open(filepath, "wb") as f:
                    f.write(payload)

                downloaded.append(saved_filename)

    finally:
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


def _send_email_smtp(to_email: str, subject: str, body: str):
    email_address, app_password = _get_gmail_credentials()

    if not to_email:
        raise ValueError("Candidatul nu are email extras din CV.")

    message = MIMEMultipart()
    message["From"] = email_address
    message["To"] = to_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain", "utf-8"))

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(email_address, app_password)
        server.send_message(message)


def send_interview_email(to_email, candidate_name, job_title):
    safe_name = candidate_name or "candidat"
    safe_job = job_title or "postul pentru care ati aplicat"

    subject = "Continuarea procesului de recrutare"

    body = f"""Buna ziua, {safe_name},

Va multumim pentru CV-ul transmis pentru postul de {safe_job}.

In urma analizei initiale, dorim sa continuam procesul de recrutare.

Un membru al echipei noastre va va contacta in urmatoarele 2 zile lucratoare pentru o prima discutie telefonica.

Va multumim,
Echipa HR
"""

    _send_email_smtp(to_email, subject, body)

    return {
        "sent": True,
        "to": to_email,
        "subject": subject
    }


def send_rejection_email(to_email, candidate_name, job_title):
    safe_name = candidate_name or "candidat"
    safe_job = job_title or "postul pentru care ati aplicat"

    subject = "Raspuns privind procesul de recrutare"

    body = f"""Buna ziua, {safe_name},

Va multumim pentru CV-ul transmis si pentru interesul acordat postului de {safe_job}.

In urma analizei initiale, am decis sa continuam procesul de recrutare cu alti candidati ale caror experiente sunt mai apropiate de cerintele actuale ale rolului.

Vom pastra datele dumneavoastra pentru oportunitati viitoare potrivite profilului transmis.

Va multumim pentru timpul acordat si va dorim mult succes in continuare.

Cu respect,
Echipa HR
"""

    _send_email_smtp(to_email, subject, body)

    return {
        "sent": True,
        "to": to_email,
        "subject": subject
    }


if __name__ == "__main__":
    result = download_cv_attachments()
    print("CV-uri descarcate:", result["downloaded_count"])

    for file in result["downloaded"]:
        print("Descarcat:", file)

    print("CV-uri ignorate/deja existente:", result["skipped_count"])
