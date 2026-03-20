#!/usr/bin/env python3
"""
Bank Branch Opportunity Scoring
================================
Uruchomienie: kliknij Run (trójkąt) w VS Code lub wciśnij F5.
Wyniki trafią do pliku wskazanego w SCIEZKA_OUTPUT.
"""

import warnings
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")


# ══════════════════════════════════════════════════════════════════════════════
# ŚCIEŻKI DO PLIKÓW — zmień tylko tutaj
# ══════════════════════════════════════════════════════════════════════════════

SCIEZKA_SIEC   = r"C:\Users\TwojeImie\dane\Struktura_sieci.xlsx"
SCIEZKA_GMINY  = r"C:\Users\TwojeImie\dane\Dane_o_gminach.csv"
SCIEZKA_OUTPUT = r"C:\Users\TwojeImie\dane\scoring_dashboard.xlsx"


# ══════════════════════════════════════════════════════════════════════════════
# WAGI SCORINGU — zmień jeśli chcesz inne proporcje (nie muszą sumować się do 1)
# ══════════════════════════════════════════════════════════════════════════════

WAGI = {
    "penetracja_gap":      0.35,   # niezagospodarowany potencjał rynkowy
    "luka_konkurencyjna":  0.25,   # udział własny vs. konkurencja
    "dynamika_wzrostu":    0.25,   # saldo migracji + przyrost naturalny
    "shift_demograficzny": 0.15,   # zmiana grupy 20-35 wg powiatu (2035)
}


# ══════════════════════════════════════════════════════════════════════════════
# NAZWY KOLUMN — zmień jeśli w Twoich plikach kolumny nazywają się inaczej
# ══════════════════════════════════════════════════════════════════════════════

COL_TERYT_GMINA  = "teryt_gmina"
COL_TERYT_POWIAT = "teryt_powiat"
COL_GMINA        = "gmina"
COL_POWIAT       = "powiat"

COL_LUDNOSC    = "Liczba mieszkańców"
COL_UDZIAL     = "udział klientów w mieszkańcach (%)"
COL_MEDIANA    = "mediana wynagrodzenia brutto"
COL_PRZYROST   = "przyrost naturalny na 1 tys mieszkańców"
COL_MIGRACJA   = "saldo migracji na 1 tys. mieszkańców"
COL_DELTA_2035 = "różnica między liczbą mieszkańców w wieku 20-35 wg powiatu"
COL_OWN        = "Liczba placówek własnych (moje)"
COL_COMP       = "Liczba placówek własnych konkurencji"


# ══════════════════════════════════════════════════════════════════════════════
# PARAMETRY MODELU
# ══════════════════════════════════════════════════════════════════════════════

BLEND_GMINA     = 0.60   # waga danych gminnych dla gmin będących częścią powiatu
BLEND_POWIAT    = 0.40   # waga danych powiatowych j.w.
MAX_PENETRATION = 0.35   # zakładany maksymalny udział klientów w populacji gminy

_BLEND_COLS = [COL_UDZIAL, COL_MEDIANA, COL_PRZYROST, COL_MIGRACJA]
_SUM_COLS   = [COL_LUDNOSC, COL_OWN, COL_COMP]

_WAGA_CELLS = {
    "penetracja_gap":      "Dashboard!$B$5",
    "luka_konkurencyjna":  "Dashboard!$B$6",
    "dynamika_wzrostu":    "Dashboard!$B$7",
    "shift_demograficzny": "Dashboard!$B$8",
}
_WAGA_SUM = "SUM(Dashboard!$B$5:Dashboard!$B$8)"
_COMP_COL_NAMES = {
    "penetracja_gap":      "c_penetracja_gap",
    "luka_konkurencyjna":  "c_luka_konkurencyjna",
    "dynamika_wzrostu":    "c_dynamika_wzrostu",
    "shift_demograficzny": "c_shift_demograficzny",
}


# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE SCORINGOWY
# ══════════════════════════════════════════════════════════════════════════════

def run_scoring(miejscowosci_df, dane_gminy_df, weights=None):
    w = weights if weights is not None else WAGI.copy()
    total = sum(w.values())
    w = {k: v / total for k, v in w.items()}   # normalizacja

    siec    = _prepare_network(miejscowosci_df)
    gminy   = _prepare_demographics(dane_gminy_df)
    gminy   = _detect_city_county(gminy)
    pow_agg = _aggregate_powiat(gminy)
    blended = _blend(gminy, pow_agg)

    result = blended.merge(
        siec[[COL_TERYT_GMINA, "n_placowek_wlasnych"]],
        on=COL_TERYT_GMINA, how="inner",
    )
    result = _compute_score_components(result)
    result = _compute_composite_score(result, w)

    return (
        result[_build_col_order(result)]
        .sort_values("opportunity_score", ascending=False)
        .reset_index(drop=True)
    )


def _prepare_network(df):
    d = df.copy()
    d.columns = d.columns.str.strip()
    d = d.rename(columns={
        _find_col(d, ["teryt_gmina", "Teryt_gmina", "TERYT_GMINA"]): COL_TERYT_GMINA,
        _find_col(d, ["gmina",       "Gmina",       "GMINA"]):        COL_GMINA,
        _find_col(d, ["powiat",      "Powiat",      "POWIAT"]):       COL_POWIAT,
    })
    return (
        d.groupby([COL_TERYT_GMINA, COL_GMINA, COL_POWIAT], as_index=False)
         .size()
         .rename(columns={"size": "n_placowek_wlasnych"})
    )


def _prepare_demographics(df):
    d = df.copy()
    d.columns = d.columns.str.strip()
    if COL_UDZIAL in d.columns and d[COL_UDZIAL].max() > 1:
        d[COL_UDZIAL] = d[COL_UDZIAL] / 100
    return d


def _detect_city_county(df):
    d = df.copy()
    name_match = (
        d[COL_GMINA].str.lower().str.strip() ==
        d[COL_POWIAT].str.lower().str.strip()
    )
    powiat_cnt  = d.groupby(COL_TERYT_POWIAT)[COL_TERYT_GMINA].transform("count")
    teryt_match = (
        (d[COL_TERYT_GMINA].str[:4] == d[COL_TERYT_POWIAT].str[:4]) &
        (powiat_cnt == 1)
    )
    d["gmina_jest_powiatem"] = name_match | teryt_match
    return d


def _aggregate_powiat(df):
    records = []
    for powiat_id, grp in df.groupby(COL_TERYT_POWIAT):
        rec = {COL_TERYT_POWIAT: powiat_id}
        pop = grp[COL_LUDNOSC].fillna(0) if COL_LUDNOSC in grp.columns \
              else pd.Series(1, index=grp.index)
        for col in _BLEND_COLS:
            if col not in grp.columns:
                continue
            valid = grp[col].notna()
            if valid.any() and pop[valid].sum() > 0:
                rec[f"p_{col}"] = np.average(grp.loc[valid, col], weights=pop[valid])
            else:
                rec[f"p_{col}"] = grp[col].mean()
        for col in _SUM_COLS:
            if col in grp.columns:
                rec[f"p_{col}"] = grp[col].fillna(0).sum()
        if COL_DELTA_2035 in grp.columns:
            rec[f"p_{COL_DELTA_2035}"] = grp[COL_DELTA_2035].iloc[0]
        records.append(rec)
    return pd.DataFrame(records)


def _blend(df_gminy, df_pow):
    d    = df_gminy.merge(df_pow, on=COL_TERYT_POWIAT, how="left")
    city = d["gmina_jest_powiatem"]

    def blend_col(col):
        gm = d[col].fillna(0)
        pw = d[f"p_{col}"].fillna(0) if f"p_{col}" in d.columns else gm
        return pd.Series(
            np.where(city, gm, BLEND_GMINA * gm + BLEND_POWIAT * pw),
            index=d.index,
        )

    for col in _BLEND_COLS + _SUM_COLS:
        if col in d.columns:
            d[f"b_{col}"] = blend_col(col)

    p_delta = f"p_{COL_DELTA_2035}"
    if p_delta in d.columns:
        d[f"b_{COL_DELTA_2035}"] = d[p_delta]
    return d


def _compute_score_components(df):
    d = df.copy()
    b = lambda col: f"b_{col}"

    if b(COL_UDZIAL) in d.columns:
        gap = (MAX_PENETRATION - d[b(COL_UDZIAL)]).clip(lower=0) / MAX_PENETRATION
        d["c_penetracja_gap"] = _minmax(gap)
    else:
        d["c_penetracja_gap"] = np.nan

    if b(COL_OWN) in d.columns and b(COL_COMP) in d.columns:
        total = d[b(COL_OWN)].fillna(0) + d[b(COL_COMP)].fillna(0)
        share = np.where(total > 0, d[b(COL_OWN)].fillna(0) / total, 0.5)
        d["c_share_oddzialow"]    = share
        d["c_luka_konkurencyjna"] = _minmax(pd.Series(1 - share, index=d.index))
    else:
        d["c_luka_konkurencyjna"] = np.nan

    parts = []
    if b(COL_MIGRACJA) in d.columns:
        parts.append((d[b(COL_MIGRACJA)].fillna(0), 0.6))
    if b(COL_PRZYROST) in d.columns:
        parts.append((d[b(COL_PRZYROST)].fillna(0), 0.4))
    if parts:
        raw = sum(s * w for s, w in parts)
        d["c_dynamika_wzrostu_raw"] = raw
        d["c_dynamika_wzrostu"]     = _minmax(raw)
    else:
        d["c_dynamika_wzrostu"] = np.nan

    if b(COL_DELTA_2035) in d.columns:
        d["c_shift_demograficzny"] = _minmax(d[b(COL_DELTA_2035)].fillna(0))
    else:
        d["c_shift_demograficzny"] = np.nan

    return d


def _compute_composite_score(df, weights):
    d = df.copy()
    comp_map = {
        "penetracja_gap":      "c_penetracja_gap",
        "luka_konkurencyjna":  "c_luka_konkurencyjna",
        "dynamika_wzrostu":    "c_dynamika_wzrostu",
        "shift_demograficzny": "c_shift_demograficzny",
    }
    active = {
        k: v for k, v in weights.items()
        if comp_map.get(k) in d.columns and d[comp_map[k]].notna().any()
    }
    norm_w = {k: v / sum(active.values()) for k, v in active.items()}
    score  = sum(norm_w[k] * d[comp_map[k]].fillna(0) for k in norm_w)
    d["opportunity_score"] = (score * 100).round(2)
    d["segment"] = pd.cut(
        d["opportunity_score"],
        bins=[0, 30, 50, 70, 100],
        labels=["Niski", "Średni", "Wysoki", "Priorytet"],
        include_lowest=True,
    )
    d["rank"] = d["opportunity_score"].rank(ascending=False, method="min").astype(int)
    return d


def _minmax(s):
    mn, mx = s.min(), s.max()
    return pd.Series(0.5, index=s.index) if mx == mn else (s - mn) / (mx - mn)


def _find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    raise KeyError(
        f"Nie znaleziono żadnej z kolumn: {candidates}\n"
        f"Dostępne kolumny: {list(df.columns)}"
    )


def _build_col_order(df):
    priority = [
        "rank", "opportunity_score", "segment",
        COL_TERYT_GMINA, COL_TERYT_POWIAT, COL_GMINA, COL_POWIAT,
        "gmina_jest_powiatem", "n_placowek_wlasnych",
        f"b_{COL_LUDNOSC}", f"b_{COL_UDZIAL}", f"b_{COL_MEDIANA}",
        f"b_{COL_PRZYROST}", f"b_{COL_MIGRACJA}", f"b_{COL_DELTA_2035}",
        f"b_{COL_OWN}", f"b_{COL_COMP}",
        "c_penetracja_gap", "c_luka_konkurencyjna",
        "c_dynamika_wzrostu_raw", "c_dynamika_wzrostu",
        "c_share_oddzialow", "c_shift_demograficzny",
    ]
    existing  = [c for c in priority if c in df.columns]
    remaining = [c for c in df.columns if c not in existing]
    return existing + remaining


# ══════════════════════════════════════════════════════════════════════════════
# EKSPORT DO EXCELA
# ══════════════════════════════════════════════════════════════════════════════

_FNT  = "Arial"
_NAVY = "1F3864"; _BMID = "2F5597"; _WHITE = "FFFFFF"; _BLACK = "000000"
_ALT  = "EEF2F7"; _HDBG = "D6E4F0"; _YEL   = "FFF2CC"; _BLUE  = "0000FF"

_SEG_STYLE = {
    "Priorytet": ("1A6B3C", _WHITE),
    "Wysoki":    ("2E86AB", _WHITE),
    "Średni":    ("F4A261", _BLACK),
    "Niski":     ("C9453A", _WHITE),
}

_THIN = Side(style="thin", color="B8C4D0")
_BRD  = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _f(bold=False, size=10, color=_BLACK):
    return Font(bold=bold, size=size, color=color, name=_FNT)

def _fill(c):
    return PatternFill("solid", fgColor=c)

def _al(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def _cell(ws, row, col, val, bold=False, size=10, color=_BLACK,
          bg=None, fmt=None, align="center", wrap=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _f(bold, size, color)
    c.alignment = _al(align, wrap)
    c.border = _BRD
    if bg:  c.fill = _fill(bg)
    if fmt: c.number_format = fmt
    return c

def _header_row(ws, row, labels, bg=_NAVY, height=32):
    for col, lbl in enumerate(labels, 1):
        c = ws.cell(row=row, column=col, value=lbl)
        c.font = _f(True, 10, _WHITE)
        c.fill = _fill(bg)
        c.alignment = _al("center", wrap=True)
        c.border = _BRD
    ws.row_dimensions[row].height = height

def _section(ws, row, c1, c2, text, bg=_BMID):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=text)
    c.font = _f(True, 11, _WHITE)
    c.fill = _fill(bg)
    c.alignment = _al("center")
    c.border = _BRD
    ws.row_dimensions[row].height = 24

def _val(v):
    if isinstance(v, (np.integer, np.int64)):    return int(v)
    if isinstance(v, (np.floating, np.float64)): return float(v)
    return v


def _build_dashboard(wb, df, weights):
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:M1")
    t = ws.cell(row=1, column=1, value="Bank · Opportunity Scoring · Dashboard")
    t.font = _f(True, 14, _WHITE); t.fill = _fill(_NAVY)
    t.alignment = _al(); ws.row_dimensions[1].height = 36

    _section(ws, 3, 1, 4, "Wagi scoringu  (żółte komórki = edytowalne)")
    _header_row(ws, 4, ["Składowa", "Waga", "% sumy", "Opis"], bg=_HDBG, height=28)
    for col in range(1, 5):
        ws.cell(row=4, column=col).font = _f(True, 10, _BLACK)

    weight_meta = [
        ("penetracja_gap",      weights["penetracja_gap"],      "Niezagospodarowany potencjał rynkowy"),
        ("luka_konkurencyjna",  weights["luka_konkurencyjna"],   "Udział własny vs. konkurencja"),
        ("dynamika_wzrostu",    weights["dynamika_wzrostu"],     "Saldo migracji 60% + przyrost 40%"),
        ("shift_demograficzny", weights["shift_demograficzny"],  "Zmiana grupy 20-35 wg powiatu"),
    ]
    for i, (key, waga, opis) in enumerate(weight_meta, start=5):
        _cell(ws, i, 1, key, align="left")
        cw = ws.cell(row=i, column=2, value=waga)
        cw.font = _f(True, 11, _BLUE); cw.fill = _fill(_YEL)
        cw.alignment = _al(); cw.border = _BRD; cw.number_format = "0.00"
        _cell(ws, i, 3, f"=B{i}/SUM($B$5:$B$8)", fmt="0.0%")
        _cell(ws, i, 4, opis, align="left")

    _cell(ws, 9, 1, "Suma kontrolna", bold=True, bg=_HDBG)
    _cell(ws, 9, 2, "=SUM(B5:B8)", bold=True, fmt="0.00")
    _cell(ws, 9, 3, "")
    _cell(ws, 9, 4, "Musi wynosić 1.00", align="left", color="7F7F7F", size=9)

    _section(ws, 11, 1, 4, "Parametry modelu")
    _header_row(ws, 12, ["Parametr", "Wartość", "Opis", ""], bg=_HDBG, height=24)
    for col in range(1, 5):
        ws.cell(row=12, column=col).font = _f(True, 10, _BLACK)
    for i, (name, val, desc) in enumerate([
        ("MAX_PENETRATION", MAX_PENETRATION, "Maks. zakładany udział klientów w populacji"),
        ("BLEND_GMINA",     BLEND_GMINA,     "Waga danych gminnych (gmina ≠ powiat)"),
        ("BLEND_POWIAT",    BLEND_POWIAT,    "Waga danych powiatowych (gmina ≠ powiat)"),
    ], start=13):
        _cell(ws, i, 1, name, align="left")
        cp = ws.cell(row=i, column=2, value=val)
        cp.font = _f(True, 11, _BLUE); cp.fill = _fill(_YEL)
        cp.alignment = _al(); cp.border = _BRD; cp.number_format = "0.00"
        _cell(ws, i, 3, desc, align="left"); _cell(ws, i, 4, "")

    _section(ws, 3, 6, 9, "Podsumowanie")
    for i, (lbl, val, color) in enumerate([
        ("Gmin w scoringu",       len(df),                                    None),
        ("Śr. Opportunity Score", round(float(df["opportunity_score"].mean()), 1), None),
        ("Segment Priorytet",     int((df["segment"] == "Priorytet").sum()),  "1A6B3C"),
        ("Segment Wysoki",        int((df["segment"] == "Wysoki").sum()),     "2E86AB"),
        ("Segment Średni",        int((df["segment"] == "Średni").sum()),     "F4A261"),
        ("Segment Niski",         int((df["segment"] == "Niski").sum()),      "C9453A"),
    ], start=4):
        bg = color or _HDBG
        tc = _WHITE if color else _BLACK
        _cell(ws, i, 6, lbl, align="left",   bg=bg, color=tc, bold=bool(color))
        _cell(ws, i, 7, val, align="center", bg=bg, color=tc, bold=True, size=12)
        ws.merge_cells(start_row=i, start_column=8, end_row=i, end_column=9)
        _cell(ws, i, 8, "")

    top10 = df.nlargest(10, "opportunity_score").reset_index(drop=True)
    _section(ws, 11, 6, 11, "TOP 10 gmin")
    _header_row(ws, 12, ["Rank", "Gmina", "Powiat", "Score", "Segment", "Placówki"],
                bg=_BMID, height=28)
    for i, row in top10.iterrows():
        r   = 13 + i
        bg  = _ALT if i % 2 else _WHITE
        seg = str(row.get("segment", ""))
        sbg, stc = _SEG_STYLE.get(seg, (bg, _BLACK))
        _cell(ws, r,  6, int(row["rank"]),                        bg=bg)
        _cell(ws, r,  7, str(row.get(COL_GMINA, "")),             bg=bg, align="left")
        _cell(ws, r,  8, str(row.get(COL_POWIAT, "")),            bg=bg, align="left")
        _cell(ws, r,  9, _val(row["opportunity_score"]),          bg=bg, fmt="0.0")
        _cell(ws, r, 10, seg,                                     bg=sbg, color=stc, bold=True)
        _cell(ws, r, 11, int(row.get("n_placowek_wlasnych", 0)), bg=bg)

    ws.cell(row=3, column=14, value="Gmina")
    ws.cell(row=3, column=15, value="Score")
    for i, row in top10.iterrows():
        ws.cell(row=4 + i, column=14, value=str(row.get(COL_GMINA, "")))
        ws.cell(row=4 + i, column=15, value=_val(row["opportunity_score"]))

    chart = BarChart()
    chart.type = "bar"; chart.grouping = "clustered"
    chart.title = "Opportunity Score – TOP 10"
    chart.y_axis.title = "Gmina"; chart.x_axis.title = "Score (0–100)"
    chart.style = 10; chart.width = 18; chart.height = 13
    chart.x_axis.numFmt = "0"
    chart.x_axis.scaling.min = 0; chart.x_axis.scaling.max = 100
    dr = Reference(ws, min_col=15, min_row=3, max_row=13)
    cr = Reference(ws, min_col=14, min_row=4, max_row=13)
    chart.add_data(dr, titles_from_data=True)
    chart.set_categories(cr)
    chart.series[0].graphicalProperties.solidFill = "2F5597"
    chart.series[0].dLbls = DataLabelList()
    chart.series[0].dLbls.showVal = True
    chart.series[0].dLbls.numFmt = "0.0"
    ws.add_chart(chart, "F24")

    for col, w in {"A": 24, "B": 10, "C": 10, "D": 40, "E": 2,
                   "F": 24, "G": 10, "H": 18, "I": 10, "J": 14, "K": 10,
                   "L": 2,  "N": 20, "O": 10}.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["N"].hidden = True
    ws.column_dimensions["O"].hidden = True


def _build_ranking(wb, df):
    ws = wb.create_sheet("Ranking")
    ws.sheet_view.showGridLines = False

    static_cols = [
        (COL_GMINA,               "Gmina",            22, "@"),
        (COL_POWIAT,              "Powiat",            18, "@"),
        ("gmina_jest_powiatem",   "Miasto/powiat",      9, "@"),
        ("n_placowek_wlasnych",   "Placówki",           8, "0"),
        (f"b_{COL_LUDNOSC}",      "Ludność",           12, "#,##0"),
        (f"b_{COL_UDZIAL}",       "Udział klientów",   12, "0.0%"),
        (f"b_{COL_MEDIANA}",      "Mediana wyngr.",    13, "#,##0"),
        (f"b_{COL_PRZYROST}",     "Przyrost nat.",     10, "0.0"),
        (f"b_{COL_MIGRACJA}",     "Saldo migracji",    10, "0.0"),
        (f"b_{COL_OWN}",          "Oddz. własne",       9, "0"),
        (f"b_{COL_COMP}",         "Oddz. konk.",        9, "0"),
        ("c_penetracja_gap",      "Penetracja gap",    12, "0.000"),
        ("c_luka_konkurencyjna",  "Luka konk.",        12, "0.000"),
        ("c_dynamika_wzrostu",    "Dynamika",          12, "0.000"),
        ("c_shift_demograficzny", "Shift dem.",        12, "0.000"),
    ]
    avail      = [(c, l, w, f) for c, l, w, f in static_cols if c in df.columns]
    total_cols = len(avail) + 3

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    t = ws.cell(row=1, column=1,
                value="Ranking — Score (Excel) przeliczany formułami z wag w Dashboard")
    t.font = _f(True, 12, _WHITE); t.fill = _fill(_NAVY); t.alignment = _al()
    ws.row_dimensions[1].height = 32

    dyn_bg = "1A6B3C"
    for col, (lbl, bg) in enumerate(
        zip([l for _, l, _, _ in avail] + ["Score (Excel)", "Segment (Excel)", "Rank (Excel)"],
            [_BMID] * len(avail) + [dyn_bg, dyn_bg, dyn_bg]),
        start=1,
    ):
        c = ws.cell(row=2, column=col, value=lbl)
        c.font = _f(True, 10, _WHITE); c.fill = _fill(bg)
        c.alignment = _al("center", wrap=True); c.border = _BRD
    ws.row_dimensions[2].height = 36

    comp_col_idx = {
        col: j for j, (col, *_) in enumerate(avail, 1)
        if col in _COMP_COL_NAMES.values()
    }
    score_ci   = len(avail) + 1
    segment_ci = len(avail) + 2
    rank_ci    = len(avail) + 3
    score_ltr  = get_column_letter(score_ci)

    df_s = df.sort_values("opportunity_score", ascending=False).reset_index(drop=True)
    n    = len(df_s)

    for i, (_, row) in enumerate(df_s.iterrows()):
        r  = 3 + i
        bg = _ALT if i % 2 else _WHITE

        for j, (col, _, _, fmt) in enumerate(avail, 1):
            c = ws.cell(row=r, column=j, value=_val(row.get(col, "")))
            c.fill = _fill(bg); c.font = _f(size=10)
            c.alignment = _al(); c.border = _BRD; c.number_format = fmt

        parts = []
        for key, waga_addr in _WAGA_CELLS.items():
            ci = comp_col_idx.get(_COMP_COL_NAMES[key])
            if ci:
                parts.append(f"{waga_addr}*{get_column_letter(ci)}{r}")
        formula_score = f"=({'+'.join(parts)})/({_WAGA_SUM})*100" if parts \
                        else _val(row["opportunity_score"])

        cs = ws.cell(row=r, column=score_ci, value=formula_score)
        cs.fill = _fill(bg); cs.font = _f(True, 10)
        cs.alignment = _al(); cs.border = _BRD; cs.number_format = "0.00"

        sc = f"{score_ltr}{r}"
        csg = ws.cell(row=r, column=segment_ci,
                      value=f'=IFS({sc}>70,"Priorytet",{sc}>50,"Wysoki",{sc}>30,"Średni",TRUE,"Niski")')
        csg.fill = _fill(bg); csg.font = _f(True, 10)
        csg.alignment = _al(); csg.border = _BRD

        crk = ws.cell(row=r, column=rank_ci,
                      value=f"=RANK.EQ({sc},${score_ltr}$3:${score_ltr}${2+n},0)")
        crk.fill = _fill(bg); crk.font = _f(True, 10)
        crk.alignment = _al(); crk.border = _BRD; crk.number_format = "0"

    last_r = 2 + n
    ws.conditional_formatting.add(
        f"{score_ltr}3:{score_ltr}{last_r}",
        ColorScaleRule(
            start_type="min",      start_color="C9453A",
            mid_type="percentile", mid_value=50, mid_color="F4A261",
            end_type="max",        end_color="1A6B3C",
        ),
    )
    seg_ltr = get_column_letter(segment_ci)
    for seg_val, (bg_c, fc) in _SEG_STYLE.items():
        ws.conditional_formatting.add(
            f"{seg_ltr}3:{seg_ltr}{last_r}",
            CellIsRule(operator="equal", formula=[f'"{seg_val}"'],
                       fill=_fill(bg_c), font=_f(True, 10, fc)),
        )

    for j, (_, _, width, _) in enumerate(avail, 1):
        ws.column_dimensions[get_column_letter(j)].width = width
    ws.column_dimensions[get_column_letter(score_ci)].width   = 14
    ws.column_dimensions[get_column_letter(segment_ci)].width = 16
    ws.column_dimensions[get_column_letter(rank_ci)].width    = 10
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}2"


def _build_skladowe(wb, df):
    ws = wb.create_sheet("Składowe")
    ws.sheet_view.showGridLines = False

    comp_cols = [
        (COL_GMINA,                "Gmina",          22, "@"),
        (COL_POWIAT,               "Powiat",          16, "@"),
        ("gmina_jest_powiatem",    "Miasto/pow.",      8, "@"),
        ("c_penetracja_gap",       "Penetracja gap",  14, "0.000"),
        ("c_luka_konkurencyjna",   "Luka konk.",      14, "0.000"),
        ("c_dynamika_wzrostu_raw", "Dynamika raw",    14, "0.0"),
        ("c_dynamika_wzrostu",     "Dynamika [0-1]",  14, "0.000"),
        ("c_shift_demograficzny",  "Shift dem.",      14, "0.000"),
        ("c_share_oddzialow",      "Share oddz.",     12, "0.0%"),
        ("opportunity_score",      "Score (Python)",  12, "0.00"),
        ("segment",                "Segment",         12, "@"),
    ]
    avail = [(c, l, w, f) for c, l, w, f in comp_cols if c in df.columns]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(avail))
    t = ws.cell(row=1, column=1,
                value="Składowe scoringu [0–1] — wartości z Pythona (audyt)")
    t.font = _f(True, 12, _WHITE); t.fill = _fill(_NAVY); t.alignment = _al()
    ws.row_dimensions[1].height = 28
    _header_row(ws, 2, [l for _, l, _, _ in avail], bg=_BMID, height=32)

    df_s = df.sort_values("opportunity_score", ascending=False).reset_index(drop=True)
    for i, (_, row) in enumerate(df_s.iterrows()):
        r = 3 + i; bg = _ALT if i % 2 else _WHITE
        for j, (col, _, _, fmt) in enumerate(avail, 1):
            c = ws.cell(row=r, column=j, value=_val(row.get(col, "")))
            c.fill = _fill(bg); c.font = _f(size=10)
            c.alignment = _al(); c.border = _BRD; c.number_format = fmt

    for j, (_, _, w, _) in enumerate(avail, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A3"

    col_idx = {c: j for j, (c, *_) in enumerate(avail, 1)}
    pg = col_idx.get("c_penetracja_gap")
    dw = col_idx.get("c_dynamika_wzrostu")
    if pg and dw:
        last_r = 2 + len(df_s)
        ch = ScatterChart()
        ch.title = "Penetracja gap vs Dynamika wzrostu"
        ch.style = 10
        ch.x_axis.title = "Penetracja gap [0–1]"
        ch.y_axis.title = "Dynamika wzrostu [0–1]"
        ch.width = 16; ch.height = 12
        xv = Reference(ws, min_col=pg, min_row=3, max_row=last_r)
        yv = Reference(ws, min_col=dw, min_row=3, max_row=last_r)
        s  = Series(yv, xv, title="Gminy")
        s.marker.symbol = "circle"; s.marker.size = 5
        s.graphicalProperties.line.noFill = True
        s.marker.graphicalProperties.solidFill      = "2F5597"
        s.marker.graphicalProperties.line.solidFill = "2F5597"
        ch.series.append(s)
        ws.add_chart(ch, f"{get_column_letter(len(avail) + 2)}3")


def _build_dane(wb, df):
    ws = wb.create_sheet("Dane")
    ws.sheet_view.showGridLines = False

    b_cols   = [c for c in df.columns if c.startswith("b_")]
    id_cols  = [COL_TERYT_GMINA, COL_TERYT_POWIAT, COL_GMINA, COL_POWIAT,
                "gmina_jest_powiatem", "n_placowek_wlasnych"]
    all_cols = [c for c in id_cols if c in df.columns] + b_cols

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(all_cols))
    t = ws.cell(row=1, column=1,
                value="Dane po blendingu (b_*) — źródło dla składowych scoringu")
    t.font = _f(True, 12, _WHITE); t.fill = _fill(_NAVY); t.alignment = _al()
    ws.row_dimensions[1].height = 28
    _header_row(ws, 2, [c.replace("b_", "").replace("_", " ") for c in all_cols],
                bg=_BMID, height=32)

    df_s = df.sort_values("opportunity_score", ascending=False).reset_index(drop=True)
    for i, (_, row) in enumerate(df_s.iterrows()):
        r = 3 + i; bg = _ALT if i % 2 else _WHITE
        for j, col in enumerate(all_cols, 1):
            c = ws.cell(row=r, column=j, value=_val(row.get(col, "")))
            c.fill = _fill(bg); c.font = _f(size=9)
            c.alignment = _al(); c.border = _BRD

    for j in range(1, len(all_cols) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    ws.freeze_panes = "A3"


def export_dashboard(df, output_path, weights=None):
    w  = weights if weights is not None else WAGI.copy()
    wb = Workbook()
    print("  [1/4] Dashboard...")
    _build_dashboard(wb, df, w)
    print("  [2/4] Ranking z formułami Excel...")
    _build_ranking(wb, df)
    print("  [3/4] Składowe...")
    _build_skladowe(wb, df)
    print("  [4/4] Dane...")
    _build_dane(wb, df)
    wb.save(output_path)


# ══════════════════════════════════════════════════════════════════════════════
# URUCHOMIENIE
# ══════════════════════════════════════════════════════════════════════════════

print("=" * 55)
print("  Bank Branch Opportunity Scoring")
print("=" * 55)

# Wczytanie plików
print("\n[1/3] Wczytywanie plików...")
siec = pd.read_excel(
    SCIEZKA_SIEC,
    dtype={"teryt_gmina": str, "Teryt_gmina": str},
)
gminy = pd.read_csv(
    SCIEZKA_GMINY,
    dtype={COL_TERYT_GMINA: str, COL_TERYT_POWIAT: str},
    sep=None,
    engine="python",
)
print(f"  Siec:  {len(siec)} wierszy")
print(f"  Gminy: {len(gminy)} wierszy")

# Scoring
print("\n[2/3] Obliczanie scoringu...")
result = run_scoring(siec, gminy, WAGI)

print(f"\n  Gmin w scoringu: {len(result)}")
print(f"  Rozkład segmentów:\n{result['segment'].value_counts().to_string()}")
print(f"\n  TOP 10:")
print(
    result[["rank", COL_GMINA, COL_POWIAT, "opportunity_score", "segment"]]
    .head(10)
    .to_string(index=False)
)

# Eksport
print("\n[3/3] Generowanie pliku Excel...")
export_dashboard(result, SCIEZKA_OUTPUT, WAGI)

print(f"\n  Gotowe! Plik zapisany w:")
print(f"  {SCIEZKA_OUTPUT}")
print("=" * 55)
