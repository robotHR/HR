import os
import time

from dotenv import load_dotenv
from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

load_dotenv()

auth_router = APIRouter()
templates = Jinja2Templates(directory="app/templates")

# Rate limiting: max 5 incercari per IP, blocat 15 minute
_MAX_ATTEMPTS = 5
_LOCKOUT_SECONDS = 900  # 15 minute
_login_attempts: dict = {}  # ip -> {"count": int, "locked_until": float}


def _get_client_ip(request: Request) -> str:
    forwarded = request.headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.client.host if request.client else "unknown"


def _is_locked(ip: str) -> tuple[bool, int]:
    entry = _login_attempts.get(ip)
    if not entry:
        return False, 0
    if entry.get("locked_until", 0) > time.time():
        remaining = int(entry["locked_until"] - time.time())
        return True, remaining
    return False, 0


def get_login_username():
    return os.getenv("HR_USERNAME", "admin")


def get_login_password():
    return os.getenv("HR_PASSWORD", "admin123")


@auth_router.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    if request.session.get("logged_in"):
        return RedirectResponse(url="/", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "error": None
        }
    )


@auth_router.post("/login")
async def login_submit(
    request: Request,
    username: str = Form(""),
    password: str = Form("")
):
    ip = _get_client_ip(request)
    locked, remaining_sec = _is_locked(ip)
    if locked:
        minutes = (remaining_sec + 59) // 60
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "request": request,
                "error": f"Prea multe incercari gresite. Incearca din nou in {minutes} minut(e)."
            },
            status_code=429
        )

    expected_username = get_login_username()
    expected_password = get_login_password()

    if username.strip() == expected_username and password == expected_password:
        _login_attempts.pop(ip, None)
        request.session["logged_in"] = True
        request.session["username"] = username.strip()
        return RedirectResponse(url="/", status_code=303)

    entry = _login_attempts.setdefault(ip, {"count": 0, "locked_until": 0.0})
    entry["count"] += 1
    if entry["count"] >= _MAX_ATTEMPTS:
        entry["locked_until"] = time.time() + _LOCKOUT_SECONDS

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "error": "User sau parola incorecta."
        },
        status_code=401
    )


@auth_router.get("/logout")
async def logout(request: Request):
    request.session.clear()

    return RedirectResponse(url="/login", status_code=303)
