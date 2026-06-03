import io
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT


PAGE_W, PAGE_H = A4
MARGIN = 2 * cm

DARK = colors.HexColor("#12121a")
ACCENT = colors.HexColor("#00d4ff")
COLOR_A = colors.HexColor("#ff6b6b")
COLOR_B = colors.HexColor("#4dabf7")
COLOR_C = colors.HexColor("#69db7c")
LIGHT_GRAY = colors.HexColor("#f0f0f0")
MID_GRAY = colors.HexColor("#cccccc")


def _class_color(cls: str):
    return {"A": COLOR_A, "B": COLOR_B, "C": COLOR_C}.get(cls, MID_GRAY)


def _page_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(MARGIN, 1.2 * cm, f"Planograma Agent — Pagina {doc.page}")
    canvas.drawRightString(PAGE_W - MARGIN, 1.2 * cm, datetime.now().strftime("%d.%m.%Y"))
    canvas.restoreState()


def generate(result: dict) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=2 * cm,
    )

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle("title", parent=styles["Title"], textColor=ACCENT, fontSize=18, spaceAfter=6)
    style_h2 = ParagraphStyle("h2", parent=styles["Heading2"], textColor=DARK, fontSize=13, spaceBefore=14, spaceAfter=4)
    style_body = ParagraphStyle("body", parent=styles["Normal"], fontSize=9, leading=14)
    style_center = ParagraphStyle("center", parent=style_body, alignment=TA_CENTER)

    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    story.append(Paragraph("Raport Optimizare Planogramă", style_title))
    story.append(Paragraph(f"Generat: {datetime.now().strftime('%d.%m.%Y %H:%M')}", style_center))
    story.append(Spacer(1, 0.5 * cm))

    # ── Section 1: Statistics ─────────────────────────────────────────────────
    story.append(Paragraph("1. Statistici Generale", style_h2))
    stats = result["stats"]
    stat_data = [
        ["Indicator", "Valoare"],
        ["Total SKU-uri", str(stats["total_skus"])],
        [f"SKU Clasa A", f"{stats['a_count']} ({stats['a_pct']}%) → {stats['a_vol_pct']}% volum"],
        [f"SKU Clasa B", f"{stats['b_count']} ({stats['b_pct']}%) → {stats['b_vol_pct']}% volum"],
        [f"SKU Clasa C", f"{stats['c_count']} ({stats['c_pct']}%) → {stats['c_vol_pct']}% volum"],
        ["Comenzi/zi", str(stats["orders_per_day"])],
        ["Angajați depozit", str(stats["employees"])],
    ]
    stat_table = Table(stat_data, colWidths=[8 * cm, 9 * cm])
    stat_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), DARK),
        ("TEXTCOLOR", (0, 0), (-1, 0), ACCENT),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
        ("GRID", (0, 0), (-1, -1), 0.5, MID_GRAY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(stat_table)
    story.append(Spacer(1, 0.4 * cm))

    # ── Section 2: Claude narrative ───────────────────────────────────────────
    story.append(Paragraph("2. Analiză AI", style_h2))
    narrative = result.get("narrative") or "Analiza AI indisponibilă temporar."
    for para in narrative.split("\n"):
        if para.strip():
            story.append(Paragraph(para.strip(), style_body))
            story.append(Spacer(1, 0.15 * cm))
    story.append(Spacer(1, 0.3 * cm))

    # ── Section 3: Top 20 ─────────────────────────────────────────────────────
    story.append(Paragraph("3. Top 20 Produse după Volum", style_h2))
    recs_sorted = sorted(result["recommendations"], key=lambda x: x["volum"], reverse=True)[:20]
    top20_data = [["#", "Cod", "Produs", "Clasă", "Volum", "Score"]]
    for i, r in enumerate(recs_sorted, 1):
        top20_data.append([
            str(i), r["cod"], r["produs"][:35], r["clasa"],
            f"{int(r['volum']):,}", f"{r['score']:.1f}",
        ])

    top20_table = Table(top20_data, colWidths=[1 * cm, 2.5 * cm, 8 * cm, 1.5 * cm, 2.5 * cm, 1.5 * cm])
    style_list = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK),
        ("TEXTCOLOR", (0, 0), (-1, 0), ACCENT),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.3, MID_GRAY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]
    for i, r in enumerate(recs_sorted, 1):
        cls_col = _class_color(r["clasa"])
        style_list.append(("BACKGROUND", (3, i), (3, i), cls_col))
        style_list.append(("TEXTCOLOR", (3, i), (3, i), colors.white))
        style_list.append(("FONTNAME", (3, i), (3, i), "Helvetica-Bold"))
    top20_table.setStyle(TableStyle(style_list))
    story.append(top20_table)

    # ── Section 4: Recommendations (first 30) ─────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("4. Recomandări SKU (primele 30)", style_h2))
    rec30 = result["recommendations"][:30]
    rec_data = [["Cod", "Produs", "Clasă", "Zonă Rec.", "Acțiune", "Volum"]]
    for r in rec30:
        rec_data.append([
            r["cod"], r["produs"][:28], r["clasa"],
            r["zona_recomandata"].replace("_", " "),
            r["actiune"], f"{int(r['volum']):,}",
        ])

    rec_table = Table(rec_data, colWidths=[2 * cm, 7 * cm, 1.5 * cm, 3.5 * cm, 2.5 * cm, 2 * cm])
    rec_style = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK),
        ("TEXTCOLOR", (0, 0), (-1, 0), ACCENT),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.3, MID_GRAY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]
    for i, r in enumerate(rec30, 1):
        cls_col = _class_color(r["clasa"])
        rec_style.append(("BACKGROUND", (2, i), (2, i), cls_col))
        rec_style.append(("TEXTCOLOR", (2, i), (2, i), colors.white))
        rec_style.append(("FONTNAME", (2, i), (2, i), "Helvetica-Bold"))
    rec_table.setStyle(TableStyle(rec_style))
    story.append(rec_table)

    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    buf.seek(0)
    return buf.read()
