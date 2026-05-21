import os

from dotenv import load_dotenv
from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates

load_dotenv()

auth_router = APIRouter()
templates = Jinja2Templates(directory="app/templates")


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
    expected_username = get_login_username()
    expected_password = get_login_password()

    if username.strip() == expected_username and password == expected_password:
        request.session["logged_in"] = True
        request.session["username"] = username.strip()

        return RedirectResponse(url="/", status_code=303)

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
