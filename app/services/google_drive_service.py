
import os
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

SERVICE_ACCOUNT_FILE = "app/service_account.json"

SCOPES = ["https://www.googleapis.com/auth/drive"]

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=SCOPES
)

drive_service = build("drive", "v3", credentials=credentials)


def upload_cv_to_drive(local_file_path, filename):

    existing = drive_service.files().list(
        q=f"name='{filename}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed=false",
        fields="files(id,name)"
    ).execute()

    if existing.get("files"):
        return existing["files"][0]

    file_metadata = {
        "name": filename,
        "parents": [GOOGLE_DRIVE_FOLDER_ID]
    }

    with open(local_file_path, "rb") as f:
        file_content = f.read()

    # resumable=False => fisierul se salveaza in quota folderului (contul tau),
    # nu in quota service account-ului (care nu are storage)
    media = MediaIoBaseUpload(
        io.BytesIO(file_content),
        mimetype="application/octet-stream",
        resumable=False
    )

    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,name"
    ).execute()

    return uploaded_file


def list_cv_files():

    query = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed=false"

    results = drive_service.files().list(
        q=query,
        fields="files(id,name)"
    ).execute()

    return results.get("files", [])


def download_cv_file(file_id, destination_path):

    request = drive_service.files().get_media(fileId=file_id)

    with open(destination_path, "wb") as file:
        downloader = MediaIoBaseDownload(file, request)

        done = False
        while done is False:
            status, done = downloader.next_chunk()

    return destination_path


def stream_cv_from_drive(filename):
    """
    Cauta fisierul dupa nume in folderul Drive si returneaza
    continutul ca BytesIO + mimetype.
    Returneaza (bytes_io, mimetype) sau (None, None) daca nu e gasit.
    """
    try:
        results = drive_service.files().list(
            q=f"name='{filename}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed=false",
            fields="files(id, name, mimeType)"
        ).execute()

        files = results.get("files", [])
        if not files:
            return None, None

        file_id = files[0]["id"]

        ext = os.path.splitext(filename)[1].lower()
        if ext == ".pdf":
            mime = "application/pdf"
        elif ext == ".docx":
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif ext == ".doc":
            mime = "application/msword"
        else:
            mime = "application/octet-stream"

        request = drive_service.files().get_media(fileId=file_id)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)

        done = False
        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)
        return buffer, mime

    except Exception:
        return None, None
