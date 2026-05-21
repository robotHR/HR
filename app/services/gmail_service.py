import os
import base64
import pickle
from email.mime.text import MIMEText
from googleapiclient.discovery import build

UPLOAD_FOLDER = "app/uploads"
TOKEN_FILE = "token.pkl"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def load_gmail_service():
    if not os.path.exists(TOKEN_FILE):
        raise FileNotFoundError("Nu exista token.pkl. Ruleaza intai autentificarea Gmail.")

    with open(TOKEN_FILE, "rb") as token:
        creds = pickle.load(token)

    service = build(
        "gmail",
        "v1",
        credentials=creds
    )

    return service


def filename_is_cv(filename):
    lower_name = filename.lower()

    allowed_extensions = (
        lower_name.endswith(".pdf")
        or lower_name.endswith(".docx")
        or lower_name.endswith(".doc")
    )

    if not allowed_extensions:
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
        "receipt"
    ]

    for blocked in blocked_keywords:
        if blocked in lower_name:
            return False

    for keyword in cv_keywords:
        if keyword in lower_name:
            return True

    return False


def download_cv_attachments(max_results=20):
    service = load_gmail_service()

    results = service.users().messages().list(
        userId="me",
        maxResults=max_results
    ).execute()

    messages = results.get("messages", [])

    downloaded = []
    skipped = []

    for msg in messages:
        message = service.users().messages().get(
            userId="me",
            id=msg["id"]
        ).execute()

        payload = message.get("payload", {})
        parts = payload.get("parts", [])

        for part in parts:
            filename = part.get("filename")

            if not filename:
                continue

            if not filename_is_cv(filename):
                continue

            filepath = os.path.join(UPLOAD_FOLDER, filename)

            if os.path.exists(filepath):
                skipped.append(filename)
                continue

            body = part.get("body", {})
            attachment_id = body.get("attachmentId")

            if not attachment_id:
                continue

            attachment = service.users().messages().attachments().get(
                userId="me",
                messageId=msg["id"],
                id=attachment_id
            ).execute()

            data = attachment.get("data")

            if not data:
                continue

            file_data = base64.urlsafe_b64decode(data)

            with open(filepath, "wb") as f:
                f.write(file_data)

            downloaded.append(filename)

    return {
        "downloaded": downloaded,
        "skipped": skipped,
        "downloaded_count": len(downloaded),
        "skipped_count": len(skipped)
    }


def create_message(to_email, subject, body):
    message = MIMEText(body, "plain", "utf-8")
    message["to"] = to_email
    message["subject"] = subject

    raw = base64.urlsafe_b64encode(
        message.as_bytes()
    ).decode("utf-8")

    return {
        "raw": raw
    }


def send_interview_email(to_email, candidate_name, job_title):
    if not to_email:
        raise ValueError("Candidatul nu are email extras din CV.")

    service = load_gmail_service()

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

    message = create_message(
        to_email=to_email,
        subject=subject,
        body=body
    )

    sent = service.users().messages().send(
        userId="me",
        body=message
    ).execute()

    return {
        "sent": True,
        "message_id": sent.get("id"),
        "to": to_email,
        "subject": subject
    }


def send_rejection_email(to_email, candidate_name, job_title):
    if not to_email:
        raise ValueError("Candidatul nu are email extras din CV.")

    service = load_gmail_service()

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

    message = create_message(
        to_email=to_email,
        subject=subject,
        body=body
    )

    sent = service.users().messages().send(
        userId="me",
        body=message
    ).execute()

    return {
        "sent": True,
        "message_id": sent.get("id"),
        "to": to_email,
        "subject": subject
    }


if __name__ == "__main__":
    result = download_cv_attachments()

    print("CV-uri descarcate:", result["downloaded_count"])

    for file in result["downloaded"]:
        print("Descarcat:", file)

    print("CV-uri deja existente:", result["skipped_count"])
