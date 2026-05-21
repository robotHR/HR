# NEXAS HR - Setup Client

Acest proiect este template-ul pentru un client separat.

Exemplu:
- Client Y are proiectul lui.
- Client Z are proiectul lui.
- Bazele de date, Gmailul, CV-urile si parola sunt separate.

## 1. Copiaza template-ul

Creezi un folder nou:

```bash
nexas_hr_client_y
```

Copiaza in el fisierele proiectului, dar fara date sensibile.

Nu copia:
- `.env`
- `token.pkl`
- `credentials.json`
- `nexas_hr.db`
- fisierele din `app/uploads/`

## 2. Creeaza fisierul .env

Copiaza:

```bash
.env.example
```

si redenumeste copia in:

```bash
.env
```

Completeaza:

```text
HR_USERNAME=admin
HR_PASSWORD=parola_clientului
SECRET_KEY=cheie_lunga_random
OPENROUTER_API_KEY=cheia_openrouter
OPENROUTER_MODEL=openai/gpt-4o-mini
DATABASE_URL=sqlite:///./nexas_hr.db
CLIENT_NAME=Client Y
```

Genereaza SECRET_KEY:

```bash
python -c "import secrets; print(secrets.token_urlsafe(32))"
```

## 3. Instaleaza dependintele

```bash
pip install -r requirements.txt
```

## 4. Configureaza Gmail

Pune fisierul:

```text
credentials.json
```

in folderul principal al proiectului.

Apoi ruleaza:

```bash
python -m app.services.gmail_auth
```

Se deschide browserul.

Clientul se autentifica in Gmail.

Dupa autentificare apare:

```text
token.pkl
```

## 5. Porneste aplicatia

```bash
uvicorn app.main:app --reload
```

Deschide:

```text
http://127.0.0.1:8000
```

Aplicatia te trimite la:

```text
/login
```

## 6. Test rapid

Fa test cu 2-3 CV-uri:

1. Pune CV-uri in `app/uploads/`
2. Intra in dashboard
3. Ruleaza analiza pentru un post
4. Verifica pagina `Candidati`
5. Testeaza `Trimite mail interviu`
6. Testeaza `Potrivit altor roluri`
7. Testeaza `Candidat Exclus`
8. Testeaza `Export Excel`

## 7. Regula importanta pentru GitHub

Repository-ul trebuie sa fie privat.

Nu urca niciodata:
- `.env`
- `token.pkl`
- `credentials.json`
- `nexas_hr.db`
- CV-uri reale
- date reale de candidati

Acestea sunt ignorate prin `.gitignore`.

## 8. Structura recomandata pentru fiecare client

```text
nexas_hr_client_y/
  app/
    api/
    core/
    models/
    services/
    templates/
    uploads/
      .gitkeep
  config/
    job_profiles.json
  requirements.txt
  .env.example
  .gitignore
  README_SETUP_CLIENT.md
```

## 9. Ce este separat per client

Fiecare client are separat:

```text
.env
credentials.json
token.pkl
nexas_hr.db
app/uploads/
user si parola de login
cont Gmail
```

Asa nu se amesteca datele intre clienti.
