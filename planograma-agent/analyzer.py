import pandas as pd
import numpy as np
from datetime import datetime
import logging

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)


def detect_format(xl: pd.ExcelFile) -> str:
    sheets = [s.strip() for s in xl.sheet_names]
    if "ABC Articole" in sheets:
        return "B"
    if "Vanzari_SKU" in sheets:
        return "A"
    return "UNKNOWN"


def parse_format_a(xl: pd.ExcelFile) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    sales = xl.parse("Vanzari_SKU")
    sales.columns = [c.strip() for c in sales.columns]

    shelf = None
    if "Configuratie_Raft" in xl.sheet_names:
        shelf = xl.parse("Configuratie_Raft")
        shelf.columns = [c.strip() for c in shelf.columns]

    premises = {}
    return sales, shelf, premises


def parse_format_b(xl: pd.ExcelFile) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    abc = xl.parse("ABC Articole")
    abc.columns = [c.strip() for c in abc.columns]

    col_map = {
        "Cod Produs": "Cod_Produs",
        "Denumire Produs": "Denumire_Produs",
        "Total Iesiri": "Total_Iesiri",
        "% din Total": "Pct_din_Total",
        "% Cumulativ": "Pct_Cumulativ",
        "Clasa ABC": "Clasa_ABC",
    }
    abc.rename(columns={k: v for k, v in col_map.items() if k in abc.columns}, inplace=True)

    premises = {}
    if "Premise si normari" in xl.sheet_names:
        prem = xl.parse("Premise si normari", header=None)
        for _, row in prem.iterrows():
            if pd.notna(row.iloc[0]) and len(row) > 1 and pd.notna(row.iloc[1]):
                key = str(row.iloc[0]).strip()
                val = row.iloc[1]
                premises[key] = val

    layout = None
    if "Layot depozit" in xl.sheet_names:
        layout = xl.parse("Layot depozit")
        layout.columns = [c.strip() for c in layout.columns]

    return abc, layout, premises


def classify_abc(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Total_Iesiri"] = pd.to_numeric(df["Total_Iesiri"], errors="coerce").fillna(0)
    df.sort_values("Total_Iesiri", ascending=False, inplace=True)
    df.reset_index(drop=True, inplace=True)

    total = df["Total_Iesiri"].sum()
    if total == 0:
        df["Clasa_ABC"] = "C"
        return df

    df["Pct_Cumulativ"] = df["Total_Iesiri"].cumsum() / total * 100

    def _class(pct):
        if pct <= 80:
            return "A"
        elif pct <= 95:
            return "B"
        return "C"

    df["Clasa_ABC"] = df["Pct_Cumulativ"].apply(_class)
    return df


def compute_metrics(df: pd.DataFrame, shelf: pd.DataFrame | None) -> pd.DataFrame:
    df = df.copy()
    df["Total_Iesiri"] = pd.to_numeric(df["Total_Iesiri"], errors="coerce").fillna(0)

    has_price = "Pret_Vanzare" in df.columns and "Cost_Produs" in df.columns
    has_faces = shelf is not None and "Fete_Alocate" in (shelf.columns if shelf is not None else [])

    if shelf is not None and "Cod_Produs" in shelf.columns:
        shelf_clean = shelf[["Cod_Produs", "Fete_Alocate"]].copy() if "Fete_Alocate" in shelf.columns else shelf[["Cod_Produs"]].copy()
        df = df.merge(shelf_clean, on="Cod_Produs", how="left")
        if "Fete_Alocate" in df.columns:
            df["Fete_Alocate"] = pd.to_numeric(df["Fete_Alocate"], errors="coerce").fillna(1).replace(0, 1)
        if "Zona" in shelf.columns:
            df = df.merge(shelf[["Cod_Produs", "Zona"]].drop_duplicates("Cod_Produs"), on="Cod_Produs", how="left")

    if "Fete_Alocate" not in df.columns:
        df["Fete_Alocate"] = 1

    df["volume_per_face"] = df["Total_Iesiri"] / df["Fete_Alocate"]

    if has_price:
        df["Pret_Vanzare"] = pd.to_numeric(df["Pret_Vanzare"], errors="coerce").fillna(0)
        df["Cost_Produs"] = pd.to_numeric(df["Cost_Produs"], errors="coerce").fillna(0)
        df["Marja_RON"] = (df["Pret_Vanzare"] - df["Cost_Produs"]) * df["Total_Iesiri"]
        df["margin_per_face"] = df["Marja_RON"] / df["Fete_Alocate"]
        vol_norm = _normalize(df["volume_per_face"])
        mar_norm = _normalize(df["margin_per_face"])
        df["global_score"] = 0.70 * vol_norm + 0.30 * mar_norm
    else:
        df["global_score"] = _normalize(df["volume_per_face"])

    avg_score = df["global_score"].mean()

    def _status(s):
        if avg_score == 0:
            return "OK"
        ratio = s / avg_score
        if ratio > 1.5:
            return "STAR"
        elif ratio >= 0.75:
            return "OK"
        return "SLAB"

    df["status"] = df["global_score"].apply(_status)
    return df


def _normalize(series: pd.Series) -> pd.Series:
    mn, mx = series.min(), series.max()
    if mx == mn:
        return pd.Series([100.0] * len(series), index=series.index)
    return (series - mn) / (mx - mn) * 100


def assign_zones(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    zone_map = {
        "A": "Picking_Fata",
        "B": "Picking_Mijloc",
        "C": "Depozitare_Spate",
    }
    df["Zona_Recomandata"] = df["Clasa_ABC"].map(zone_map).fillna("Depozitare_Spate")

    def _action(row):
        clasa = row.get("Clasa_ABC", "C")
        status = row.get("status", "OK")
        zona_curenta = str(row.get("Zona", "")).strip()
        zona_rec = row.get("Zona_Recomandata", "")
        faces = row.get("Fete_Alocate", 1)

        if clasa == "A" and zona_curenta and zona_curenta not in ("Picking_Fata", ""):
            return "MUTA_FATA"
        if clasa == "C" and faces > 2:
            return "REDUCE_FETE"
        if (clasa == "A" or status == "STAR") and faces < 2:
            return "CRESTE_FETE"
        return "OK"

    df["Actiune"] = df.apply(_action, axis=1)
    return df


def build_stats(df: pd.DataFrame, premises: dict) -> dict:
    total = len(df)
    counts = df["Clasa_ABC"].value_counts().to_dict()
    a_count = counts.get("A", 0)
    b_count = counts.get("B", 0)
    c_count = counts.get("C", 0)

    total_vol = df["Total_Iesiri"].sum()

    def vol_pct(cls):
        v = df[df["Clasa_ABC"] == cls]["Total_Iesiri"].sum()
        return round(v / total_vol * 100, 1) if total_vol else 0

    orders_per_day = _extract_premise(premises, ["Comenzi medii zilnice", "comenzi", "orders"])
    employees = _extract_premise(premises, ["Număr angajați în depozit", "angajati", "employees"])

    top5 = df.nlargest(5, "Total_Iesiri")[["Cod_Produs", "Denumire_Produs", "Total_Iesiri", "Clasa_ABC"]].to_dict("records")

    problems = []
    wrong_zone = df[(df["Clasa_ABC"] == "A") & (df["Actiune"] == "MUTA_FATA")]
    if not wrong_zone.empty:
        problems.append(f"{len(wrong_zone)} produse Clasa A plasate în zonă greșită")
    reduce = df[df["Actiune"] == "REDUCE_FETE"]
    if not reduce.empty:
        problems.append(f"{len(reduce)} produse Clasa C cu prea multe fețe alocate")
    increase = df[df["Actiune"] == "CRESTE_FETE"]
    if not increase.empty:
        problems.append(f"{len(increase)} produse performante (A/STAR) cu prea puține fețe")

    return {
        "total_skus": total,
        "a_count": a_count,
        "b_count": b_count,
        "c_count": c_count,
        "a_pct": round(a_count / total * 100, 1) if total else 0,
        "b_pct": round(b_count / total * 100, 1) if total else 0,
        "c_pct": round(c_count / total * 100, 1) if total else 0,
        "a_vol_pct": vol_pct("A"),
        "b_vol_pct": vol_pct("B"),
        "c_vol_pct": vol_pct("C"),
        "orders_per_day": orders_per_day,
        "employees": employees,
        "top5": top5,
        "problems": problems,
    }


def _extract_premise(premises: dict, keys: list) -> str:
    for k in keys:
        for pk, pv in premises.items():
            if k.lower() in pk.lower():
                return str(pv)
    return "N/A"


def build_chart_data(df: pd.DataFrame) -> dict:
    counts = df["Clasa_ABC"].value_counts().to_dict()
    total = len(df)
    total_vol = df["Total_Iesiri"].sum()

    classes = ["A", "B", "C"]
    sku_pcts = []
    vol_pcts = []
    for c in classes:
        cnt = counts.get(c, 0)
        sku_pcts.append(round(cnt / total * 100, 1) if total else 0)
        vol = df[df["Clasa_ABC"] == c]["Total_Iesiri"].sum()
        vol_pcts.append(round(vol / total_vol * 100, 1) if total_vol else 0)

    return {
        "labels": classes,
        "sku_pcts": sku_pcts,
        "vol_pcts": vol_pcts,
    }


def build_recommendations(df: pd.DataFrame) -> list:
    recs = []
    for _, row in df.iterrows():
        recs.append({
            "cod": str(row.get("Cod_Produs", "")),
            "produs": str(row.get("Denumire_Produs", ""))[:60],
            "clasa": str(row.get("Clasa_ABC", "")),
            "zona_curenta": str(row.get("Zona", "N/A")),
            "zona_recomandata": str(row.get("Zona_Recomandata", "")),
            "actiune": str(row.get("Actiune", "OK")),
            "volum": float(row.get("Total_Iesiri", 0)),
            "score": round(float(row.get("global_score", 0)), 2),
            "status": str(row.get("status", "OK")),
        })
    return recs


def analyze(xl: pd.ExcelFile) -> dict:
    fmt = detect_format(xl)
    log.info("Detected format: %s", fmt)

    if fmt == "A":
        sales, shelf, premises = parse_format_a(xl)
        if "Clasa_ABC" not in sales.columns:
            sales = classify_abc(sales)
        df = compute_metrics(sales, shelf)
        df = assign_zones(df)
    elif fmt == "B":
        abc, layout, premises = parse_format_b(xl)
        if "Clasa_ABC" not in abc.columns:
            abc = classify_abc(abc)
        df = compute_metrics(abc, None)
        df = assign_zones(df)
    else:
        sheets = xl.sheet_names
        raise ValueError(
            f"Format nerecunoscut. Foi găsite: {sheets}. "
            "Așteptat: 'Vanzari_SKU' (Format A) sau 'ABC Articole' (Format B)."
        )

    stats = build_stats(df, premises)
    chart_data = build_chart_data(df)
    recommendations = build_recommendations(df)

    return {
        "stats": stats,
        "chart_data": chart_data,
        "recommendations": recommendations,
        "df": df,
        "premises": premises,
    }
