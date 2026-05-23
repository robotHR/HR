import os
import requests
from io import BytesIO

import cloudinary
import cloudinary.uploader
import cloudinary.api

cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True
)


def upload_cv_to_cloudinary(local_file_path, filename):
    """
    Uploadeaza CV pe Cloudinary ca fisier raw (PDF/DOCX).
    Daca exista deja, returneaza direct fara re-upload.
    """
    public_id = filename

    try:
        existing = cloudinary.api.resource(public_id, resource_type="raw")
        if existing:
            return existing
    except Exception:
        pass

    result = cloudinary.uploader.upload(
        local_file_path,
        public_id=public_id,
        resource_type="raw",
        overwrite=False,
        use_filename=True,
        unique_filename=False
    )
    return result


def upload_cv_bytes_to_cloudinary(file_bytes, filename):
    """
    Uploadeaza CV din bytes (ex: din Gmail) pe Cloudinary.
    """
    public_id = filename

    try:
        existing = cloudinary.api.resource(public_id, resource_type="raw")
        if existing:
            return existing
    except Exception:
        pass

    result = cloudinary.uploader.upload(
        file_bytes,
        public_id=public_id,
        resource_type="raw",
        overwrite=False,
        use_filename=True,
        unique_filename=False
    )
    return result


def stream_cv_from_cloudinary(filename):
    """
    Cauta fisierul pe Cloudinary dupa public_id (= filename)
    si returneaza (BytesIO, mimetype) sau (None, None).
    """
    try:
        resource = cloudinary.api.resource(filename, resource_type="raw")
        url = resource.get("secure_url")

        if not url:
            return None, None

        response = requests.get(url, timeout=30)
        if response.status_code != 200:
            return None, None

        ext = os.path.splitext(filename)[1].lower()
        if ext == ".pdf":
            mime = "application/pdf"
        elif ext == ".docx":
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif ext == ".doc":
            mime = "application/msword"
        else:
            mime = "application/octet-stream"

        return BytesIO(response.content), mime

    except Exception:
        return None, None
