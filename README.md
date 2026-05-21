NEXAS HR FIX

Inlocuieste fisierele in proiect exact pe aceleasi cai.

Dupa inlocuire:

CTRL + C ca sa opresti serverul

python -m app.services.cv_parser

uvicorn app.main:app --reload

Browser:
http://127.0.0.1:8000

Test DB:
http://127.0.0.1:8000/api/debug
