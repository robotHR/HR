import os
import logging
import anthropic

log = logging.getLogger(__name__)

MODEL = "claude-sonnet-4-20250514"
MAX_TOKENS = 1024


def build_prompt(stats: dict) -> str:
    top5_lines = []
    for i, p in enumerate(stats.get("top5", []), 1):
        top5_lines.append(
            f"  {i}. [{p.get('Clasa_ABC', p.get('clasa', ''))}] "
            f"{p.get('Denumire_Produs', p.get('produs', ''))} "
            f"— {int(p.get('Total_Iesiri', p.get('volum', 0)))} unități"
        )

    problems_lines = []
    for prob in stats.get("problems", []):
        problems_lines.append(f"  - {prob}")
    if not problems_lines:
        problems_lines.append("  - Nu au fost identificate probleme majore")

    return f"""Analizează următoarele date de depozit și generează un raport executiv în română:

STATISTICI GENERALE:
- Total SKU-uri: {stats['total_skus']}
- Clasa A: {stats['a_count']} SKU-uri ({stats['a_pct']}% din total) → {stats['a_vol_pct']}% din volum
- Clasa B: {stats['b_count']} SKU-uri ({stats['b_pct']}%) → {stats['b_vol_pct']}% din volum
- Clasa C: {stats['c_count']} SKU-uri ({stats['c_pct']}%) → {stats['c_vol_pct']}% din volum
- Comenzi/zi: {stats['orders_per_day']}
- Angajați depozit: {stats['employees']}

TOP 5 PRODUSE (volum):
{chr(10).join(top5_lines)}

PROBLEME IDENTIFICATE:
{chr(10).join(problems_lines)}

Generează:
1. Sumar executiv (3-4 fraze)
2. Principalele 3 concluzii din analiza ABC
3. Top 5 recomandări specifice de acțiune
4. Estimare impact dacă se aplică recomandările (timp picking, eficiență)

Răspunde în română, ton profesional, maxim 400 cuvinte."""


def get_narrative(stats: dict) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        log.warning("ANTHROPIC_API_KEY not set — skipping AI narrative")
        return ""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=(
                "Ești un expert în optimizarea depozitelor și category management. "
                "Analizezi date reale de depozit și generezi recomandări clare, practice, "
                "în limba română. Fii direct și specific — nu generaliza."
            ),
            messages=[{"role": "user", "content": build_prompt(stats)}],
        )
        return message.content[0].text
    except Exception as exc:
        log.error("Claude API error: %s", exc)
        return ""
