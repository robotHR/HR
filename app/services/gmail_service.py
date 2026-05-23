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


from app.services.google_drive_service import upload_cv_to_drive

UPLOAD_FOLDER = "app/uploads"

GMAIL_EMAIL = os.getenv("GMAIL_EMAIL")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
GMAIL_IMAP_HOST = os.getenv("GMAIL_IMAP_HOST", "imap.gmail.com")
GMAIL_IMAP_FOLDER = os.getenv("GMAIL_IMAP_FOLDER", "INBOX")

RESEND_API_KEY = os.getenv("RESEND_API_KEY")
FROM_EMAIL = os.getenv("FROM_EMAIL", "onboarding@resend.dev")

ALLOWED_EXTENSIONS = (".pdf", ".docx", ".doc")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


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


def _unique_path(folder, filename):
    base, ext = os.path.splitext(filename)
    candidate = filename
    index = 1

    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base}_{index}{ext}"
        index += 1

    return os.path.join(folder, candidate), candidate


def _attachment_already_exists(folder, filename):
    """
    Verifica daca atasamentul exista deja in uploads.
    Evita descarcarea aceluiasi CV la fiecare rulare Gmail.
    """
    base, ext = os.path.splitext(filename)

    if os.path.exists(os.path.join(folder, filename)):
        return True

    pattern = re.compile(rf"^{re.escape(base)}(_\d+)?{re.escape(ext)}$", re.IGNORECASE)

    try:
        for existing_file in os.listdir(folder):
            if pattern.match(existing_file):
                return True
    except FileNotFoundError:
        return False

    return False


def filename_is_cv(filename):
    lower_name = (filename or "").lower().strip()

    if not lower_name.endswith(ALLOWED_EXTENSIONS):
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
        "ordin",
        "oferta"
    ]

    for blocked in blocked_keywords:
        if blocked in lower_name:
            return False

    cv_keywords = [
        "cv",
        "c.v",
        "curriculum",
        "vitae",
        "resume",
        "candidat",
        "candidatura",
        "angajare",
        "aplicare"
    ]

    for keyword in cv_keywords:
        if keyword in lower_name:
            return True

    # Acceptam si fisiere PDF/DOC/DOCX fara cuvantul "CV".
    # Multi candidati trimit fisierul doar cu numele lor.
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


def _download_attachments_from_message(mail, message_id, downloaded, skipped):
    status, data = mail.fetch(message_id, "(RFC822)")

    if status != "OK" or not data or not data[0]:
        return

    raw_email = data[0][1]
    message = email.message_from_bytes(raw_email)

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

        if _attachment_already_exists(UPLOAD_FOLDER, filename):
            skipped.append(filename)
            continue

        filepath, final_filename = _unique_path(UPLOAD_FOLDER, filename)

        with open(filepath, "wb") as file:
            file.write(payload)

        upload_cv_to_drive(filepath, final_filename)

        downloaded.append(final_filename)


def download_cv_attachments(max_results=20):
    """
    Citeste ultimele emailuri din Gmail prin IMAP si descarca atasamentele CV in app/uploads.

    Necesita in Render:
    GMAIL_EMAIL
    GMAIL_APP_PASSWORD

    Optional:
    GMAIL_IMAP_FOLDER=INBOX
    """
    mail = None
    downloaded = []
    skipped = []

    try:
        mail = _connect_imap()

        # Gmail IMAP nu accepta stabil comanda: HAS ATTACHMENT.
        # Folosim X-GM-RAW, adica sintaxa Gmail reala.
        # Cauta doar emailuri recente cu atasamente, ca sa reducem timpul de procesare.
        try:
            status, data = mail.search(
                None,
                "X-GM-RAW",
                '"has:attachment newer_than:30d"'
            )
        except imaplib.IMAP4.error:
            status, data = mail.search(None, "ALL")

        if status != "OK":
            status, data = mail.search(None, "ALL")

        ids = data[0].split() if data and data[0] else []
        ids = ids[-max_results:]
        ids.reverse()

        for message_id in ids:
            _download_attachments_from_message(mail, message_id, downloaded, skipped)

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
    """
    Trimite email prin Resend API, nu SMTP.

    Necesita in Render:
    RESEND_API_KEY
    FROM_EMAIL

    Pentru test:
    FROM_EMAIL=onboarding@resend.dev

    Pentru productie:
    FROM_EMAIL=NEXAS HR <hr@domeniul-clientului.ro>
    """
    if resend is None:
        raise RuntimeError("Lipseste pachetul resend. Adauga 'resend' in requirements.txt.")

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

    result = send_email_api(
        to_email=to_email,
        subject=subject,
        body=body
    )

    return {
        "sent": True,
        "provider": "resend",
        "result": result,
        "to": to_email,
        "subject": subject
    }


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

    result = send_email_api(
        to_email=to_email,
        subject=subject,
        body=body
    )

    return {
        "sent": True,
        "provider": "resend",
        "result": result,
        "to": to_email,
        "subject": subject
    }


if __name__ == "__main__":
    result = download_cv_attachments()
    print("CV-uri descarcate:", result["downloaded_count"])
    for file in result["downloaded"]:
        print("Descarcat:", file)
    print("CV-uri ignorate/deja procesate:", result["skipped_count"])
