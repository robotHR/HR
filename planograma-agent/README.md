# Planograma Agent

Instrument inteligent de optimizare a sloturilor de depozit pentru Category Managers.

## Funcționalități

- Upload fișier Excel cu date de vânzări / configurație raft
- Clasificare ABC automată a SKU-urilor
- Calculul scorului de performanță per SKU (volum + marjă)
- Recomandări de zonare în depozit (Picking Față / Mijloc / Depozitare Spate)
- Analiză narativă generată de Claude AI (în română)
- Dashboard interactiv cu hartă vizuală a depozitului
- Export Excel și PDF

## Formate de intrare acceptate

### Format A — Standard
- Foaie `Vanzari_SKU`: Cod_Produs, Denumire_Produs, Total_Iesiri, Pret_Vanzare (opțional), Cost_Produs (opțional)
- Foaie `Configuratie_Raft`: Cod_Produs, Fete_Alocate, Zona, Nivel

### Format B — Studiu de Caz
- Foaie `ABC Articole`: Cod Produs, Denumire Produs, Total Iesiri, % din Total, % Cumulativ, Clasa ABC
- Foaie `Layot depozit` (opțional)
- Foaie `Premise si normari` (opțional)

## Instalare locală

```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env       # adăugați ANTHROPIC_API_KEY
python app.py
```

Deschideți http://localhost:5000

## Variabile de mediu

| Variabilă | Descriere | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | Cheia API Anthropic (obligatorie pentru analiza AI) | — |
| `PORT` | Portul aplicației | 5000 |
| `MAX_FILE_SIZE_MB` | Dimensiunea maximă a fișierului upload | 10 |
| `SECRET_KEY` | Cheie secretă Flask | random |

## Deploy pe Render.com

1. Push codul pe GitHub (repo privat)
2. Pe Render: **New → Web Service**
3. Conectați repo-ul GitHub
4. Setări:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. Adăugați variabila de mediu `ANTHROPIC_API_KEY` în Environment

## Structura proiectului

```
planograma-agent/
├── app.py              # Flask routes
├── analyzer.py         # Logica ABC + metrici + recomandări
├── claude_client.py    # Integrare Claude API
├── report_excel.py     # Generator raport Excel
├── report_pdf.py       # Generator raport PDF
├── templates/
│   └── index.html      # SPA frontend
├── static/
│   └── style.css       # Dark theme
└── requirements.txt
```
