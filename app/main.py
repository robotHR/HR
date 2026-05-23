import os

from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware

from app.core.database import Base, engine

from app.models.candidate_model import Candidate
from app.models.candidate_event_model import CandidateEvent
from app.models.cv_analysis_model import CvAnalysis

Base.metadata.create_all(bind=engine)

from app.api.routes import router
from app.api.auth_routes import auth_router

load_dotenv()

Base.metadata.create_all(bind=engine)

app = FastAPI()


@app.middleware("http")
async def protect_private_pages(request: Request, call_next):
    path = request.url.path

    public_paths = [
        "/login",
        "/static",
        "/favicon.ico"
    ]

    is_public = any(
        path == public_path or path.startswith(public_path + "/")
        for public_path in public_paths
    )

    if is_public:
        return await call_next(request)

    if not request.session.get("logged_in"):
        return RedirectResponse(url="/login", status_code=303)

    return await call_next(request)


SECRET_KEY = os.getenv("SECRET_KEY", "schimba-aceasta-cheie-in-env")

app.add_middleware(
    SessionMiddleware,
    secret_key=SECRET_KEY,
    max_age=60 * 60 * 8,
    same_site="lax",
    https_only=False
)

app.mount(
    "/static",
    StaticFiles(directory="app/static"),
    name="static"
)

app.include_router(auth_router)
app.include_router(router)
