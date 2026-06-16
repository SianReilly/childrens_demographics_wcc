# Westminster Children's Demographics Dashboard
# ──────────────────────────────────────────────────────────────────────────────
# pip install streamlit plotly pandas numpy openpyxl python-pptx kaleido topojson pyproj
#
# Data sources
# 1. ONS Mid-Year Population Estimates (MYEs) — LA 1991–2024 (gender) & LSOA 2022–2024
# 2. IoD 2025 — Index of Multiple Deprivation (IMD) composite rank
# 3. IoD 2025 — Supplementary Indices: IDACI (children) & IDAOPI (older people)
# 4. EGDI — Ethnic Group Deprivation Index — Lloyd et al. 2023 / gedi.ac.uk
# 5. Census 2021 — RM006 (household type), RM012 (HRP ethnicity × age), RM033 (child ethnicity × sex)
# 6. DWP Children in Low Income Families (AHC Relative) FYE 2024–2025
# 7. DfE Key Stage 4 attainment — Explore Education Statistics
# 8. Westminster ward + London borough boundaries — ONS / London Datastore
# ──────────────────────────────────────────────────────────────────────────────

import os, io, json, copy, warnings
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    _HAVE_PPTX = True
except ImportError:
    _HAVE_PPTX = False

warnings.filterwarnings("ignore")

# ── PATHS ─────────────────────────────────────────────────────────────────────
_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_DATA_SUBDIR = os.path.join(_SCRIPT_DIR, "data")
_SEARCH_DIRS = [_DATA_SUBDIR, _SCRIPT_DIR, "/mnt/user-data/uploads"]

_ALIASES = {
    "data-key-stage-4-performance__1_.csv": [
        "data-key-stage-4-performance__1_.csv", "data-key-stage-4-performance_1_.csv"],
    "ONS_LSOA_2021 (1).json": [
        "ONS_LSOA_2021 (1).json", "ONS_LSOA_2021__1_.json", "ONS_LSOA_2021_(1).json"],
    "Ward LSOA Lookup.xlsx": [
        "Ward LSOA Lookup.xlsx", "Ward_LSOA_Lookup.xlsx", "Ward LSOA Lookup .xlsx"],
}

def _dp(name):
    aliases = list(dict.fromkeys(_ALIASES.get(name, []) + [name]))
    for d in _SEARCH_DIRS:
        for a in aliases:
            p = os.path.join(d, a)
            if os.path.exists(p):
                return p
    return os.path.join("/mnt/user-data/uploads", name)

def _exists(name):
    return os.path.exists(_dp(name))

# ── WCC BRAND + ECONOMIST "GREY-THE-CONTEXT" PALETTE ──────────────────────────
WCC = {
    "blue": "#0B2265", "yellow": "#F5CB00", "cobalt": "#0C35FA",
    "amaranth": "#E34063", "green": "#008466", "orange": "#EA6F06",
    "white": "#FFFFFF", "light_blue": "#E8EBF5", "grid": "#ECECEC",
}

# Focal vs context. Westminster (the story) is in strong colour; everyone else
# is a pale/muted accent so the eye lands on Westminster. (Economist principle.)
FOCAL        = "#0B2265"   # Westminster — strong WCC blue
FOCAL_ALT    = "#0C35FA"   # secondary strong (used for ward top-3 etc.)
CONTEXT_BAR  = "#C8D0DE"   # pale grey-blue for non-focal bars
CONTEXT_LINE = "#AAB4C6"   # pale grey-blue for non-focal lines
TEXT         = "#222222"

NEIGHBOURS = {  # CIPFA statistical neighbours
    "Westminster":          "E09000033",
    "Kensington & Chelsea": "E09000020",
    "Camden":               "E09000007",
    "Hammersmith & Fulham": "E09000013",
    "Islington":            "E09000019",
    "Wandsworth":           "E09000032",
}
WARD_TOP3 = ["Westbourne", "Church Street", "Harrow Road"]  # highest child poverty

def borough_palette(categories, focal="Westminster"):
    """Strong colour for Westminster, pale grey-blue for every other borough."""
    return {c: (FOCAL if c == focal else CONTEXT_BAR) for c in categories}

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Westminster Children's Demographics",
                   page_icon="🏙️", layout="wide", initial_sidebar_state="expanded")

st.markdown(f"""<style>
  [data-testid="stSidebar"] {{ background-color:{WCC['blue']} !important; border-top:4px solid {WCC['yellow']}; }}
  [data-testid="stSidebar"] * {{ color:#fff !important; }}
  [data-testid="stSidebar"] a {{ color:#A8C0FF !important; }}
  [data-testid="stSidebar"] hr {{ border-color:rgba(255,255,255,.25) !important; }}
  h1,h2,h3 {{ color:{WCC['blue']} !important; font-family:Arial,sans-serif !important; }}
  .stTabs [data-baseweb="tab"] {{ font-size:.95rem; font-family:Arial,sans-serif; }}
  .stTabs [aria-selected="true"] {{ color:{WCC['blue']} !important;
      border-bottom:3px solid {WCC['cobalt']} !important; font-weight:700; }}
  [data-testid="stMetric"] {{ background:{WCC['light_blue']}; border-radius:6px;
      padding:12px 14px; border-left:4px solid {WCC['blue']}; }}
  .chart-title {{ font-size:1.18rem; font-weight:800; color:{WCC['blue']};
      font-family:Arial,sans-serif; margin:.2rem 0 .1rem 0; line-height:1.3; }}
  .chart-sub {{ font-size:.9rem; color:#555; margin:0 0 .35rem 0; }}
  .source-box {{ background:{WCC['light_blue']}; border-radius:5px; padding:8px 12px;
      font-size:.82em; color:#333; margin-top:6px; border-left:3px solid {WCC['blue']}; }}
  .howto {{ background:#FFF8E1; border-radius:5px; padding:7px 12px; font-size:.82em;
      color:#5d4b00; margin:4px 0 10px 0; border-left:3px solid {WCC['yellow']}; }}
  .ds-card {{ background:{WCC['light_blue']}; border-radius:6px; padding:12px 16px;
      margin-bottom:10px; border-left:4px solid {WCC['cobalt']}; }}
</style>""", unsafe_allow_html=True)

# Plotly modebar exports PNG (not SVG/HTML) so the camera icon always yields a
# slide-ready image even where server-side kaleido/Chrome is unavailable.
PLOTLY_CONFIG = {
    "displaylogo": False,
    "toImageButtonOptions": {"format": "png", "scale": 2, "filename": "westminster_chart"},
    "modeBarButtonsToRemove": ["lasso2d", "select2d"],
}

# ── PRESENTATION HELPERS ──────────────────────────────────────────────────────
def chart_title(title, subtitle=None):
    """Bold, larger chart title rendered ABOVE the chart (per house style)."""
    st.markdown(f"<div class='chart-title'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='chart-sub'>{subtitle}</div>", unsafe_allow_html=True)

def legend_hint(extra=""):
    st.markdown(
        "<div class='howto'>💡 <b>Tip:</b> click a name in the legend to hide it, "
        "or double-click a name to isolate it — handy for dropping comparators "
        f"(e.g. Camden, Hammersmith &amp; Fulham) to focus on Westminster. {extra}</div>",
        unsafe_allow_html=True)

def source_line(text):
    st.markdown(f"<div class='source-box'>{text}</div>", unsafe_allow_html=True)

def apply_style(fig, source="", height=None):
    """House styling. Titles live in chart_title() above the figure, never in plotly."""
    fig.update_layout(
        font_family="Arial", font_color=TEXT, plot_bgcolor="white", paper_bgcolor="white",
        title={"text": ""}, margin=dict(l=55, r=30, t=14, b=58),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                    font=dict(family="Arial", size=11), title=dict(text="")))
    fig.update_xaxes(showgrid=False, linecolor="#cccccc", showline=True,
                     tickfont=dict(family="Arial", size=11))
    fig.update_yaxes(gridcolor=WCC["grid"], linecolor="white", zeroline=False,
                     tickfont=dict(family="Arial", size=11))
    if height:
        fig.update_layout(height=height)
    if source:
        fig.add_annotation(text=f"<i>Source: {source}</i>", xref="paper", yref="paper",
                           x=0, y=-0.17, showarrow=False, align="left",
                           font=dict(size=9, color="#555", family="Arial"))
    return fig

def png_button(fig, key, w=1200, h=700):
    """PNG download (slide-ready). Falls back to the chart camera icon if the
    server has no Chrome/kaleido. HTML is never offered — it can't drop into PPT."""
    try:
        png = fig.to_image(format="png", width=w, height=h, scale=2)
        st.download_button("⬇ Download chart (PNG)", data=png,
                           file_name=f"{key}.png", mime="image/png", key=f"png_{key}")
    except Exception:
        st.caption("⬇ Use the 📷 camera icon on the chart toolbar to save a PNG "
                   "for slides (server-side PNG export unavailable here).")

def show_chart(fig, key, source="", height=None):
    """Standard render: style → display (PNG modebar) → source → PNG button."""
    apply_style(fig, source=source, height=height)
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG, key=f"ch_{key}")
    png_button(fig, key)

# ── GEO HELPERS ───────────────────────────────────────────────────────────────
def _inject_id(geojson, id_property):
    out = {"type": "FeatureCollection", "features": []}
    for f in geojson["features"]:
        fc = {k: v for k, v in f.items()}
        fc["id"] = f["properties"].get(id_property)
        out["features"].append(fc)
    return out

def _looks_like_osgb(geojson):
    """Detect British National Grid (eastings/northings >> 1000)."""
    try:
        c = geojson["features"][0]["geometry"]["coordinates"]
        while isinstance(c, list) and isinstance(c[0], list):
            c = c[0]
        return abs(c[0]) > 1000
    except Exception:
        return False

def _reproject_osgb_to_wgs84(geojson):
    try:
        from pyproj import Transformer
    except ImportError:
        st.warning("pyproj not installed — cannot reproject OSGB coordinates. Add 'pyproj' to requirements.txt.")
        return geojson
    tr = Transformer.from_crs("EPSG:27700", "EPSG:4326", always_xy=True)
    out = copy.deepcopy(geojson)
    for feat in out["features"]:
        g = feat["geometry"]
        if g["type"] == "Polygon":
            g["coordinates"] = [[list(tr.transform(x, y)) for x, y in ring] for ring in g["coordinates"]]
        elif g["type"] == "MultiPolygon":
            g["coordinates"] = [[[list(tr.transform(x, y)) for x, y in ring] for ring in poly]
                                for poly in g["coordinates"]]
    return out

def _ensure_wgs84(geojson):
    return _reproject_osgb_to_wgs84(geojson) if _looks_like_osgb(geojson) else geojson

def _topojson_to_geojson(topo, object_name=None):
    """Decode a TopoJSON (delta + transform) into a WGS84 GeoJSON FeatureCollection."""
    if object_name is None:
        object_name = list(topo["objects"].keys())[0]
    tr = topo.get("transform")
    def dec(arc):
        if not tr:
            return [list(p) for p in arc]
        (sx, sy), (tx, ty) = tr["scale"], tr["translate"]
        out, x, y = [], 0, 0
        for p in arc:
            x += p[0]; y += p[1]
            out.append([x * sx + tx, y * sy + ty])
        return out
    arcs = [dec(a) for a in topo["arcs"]]
    def stitch(idxs):
        coords = []
        for i in idxs:
            a = arcs[~i][::-1] if i < 0 else arcs[i]
            coords.extend(a[1:] if (coords and coords[-1] == a[0]) else a)
        return coords
    def geom(g):
        t = g["type"]
        if t == "Polygon":
            return {"type": "Polygon", "coordinates": [stitch(r) for r in g["arcs"]]}
        if t == "MultiPolygon":
            return {"type": "MultiPolygon",
                    "coordinates": [[stitch(r) for r in poly] for poly in g["arcs"]]}
        return None
    feats = [{"type": "Feature", "properties": g.get("properties", {}), "geometry": geom(g)}
             for g in topo["objects"][object_name]["geometries"]]
    return {"type": "FeatureCollection", "features": feats}

def choropleth(geojson, codes, z, names, label, colorscale, *,
               wards=None, fmt=":.1f", zoom=12, center=None, height=540, reverse=False):
    """Generic LSOA/ward/borough choropleth. `wards` (parallel list) shown on hover."""
    if center is None:
        center = {"lat": 51.515, "lon": -0.16}
    if wards is not None:
        custom = np.array(wards, dtype=object).reshape(-1, 1)
        hover = "<b>%{text}</b><br>Ward: %{customdata[0]}<br>" + label + ": %{z" + fmt + "}<extra></extra>"
    else:
        custom = None
        hover = "<b>%{text}</b><br>" + label + ": %{z" + fmt + "}<extra></extra>"
    fig = go.Figure(go.Choroplethmap(
        geojson=geojson, locations=codes, z=z, text=names, customdata=custom,
        hovertemplate=hover, colorscale=colorscale, reversescale=reverse,
        marker_opacity=0.78, marker_line_width=0.4, marker_line_color="white",
        colorbar=dict(title=dict(text=label, font=dict(size=11)), thickness=14, len=0.62)))
    fig.update_layout(map_style="carto-positron", map_zoom=zoom, map_center=center,
                      margin=dict(l=0, r=0, t=0, b=0), height=height, paper_bgcolor="white")
    return fig

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADERS — NEW uploaded datasets (parsing validated against the real files)
# ══════════════════════════════════════════════════════════════════════════════
MYE_AGE_ORDER = ["0-4", "5-9", "10-14", "15-19"]

@st.cache_data(show_spinner=False)
def load_mye_la():
    """ONS MYEs, London boroughs, 1991–2024. Stacked Nomis blocks:
    3 genders (Total/Male/Female) × 4 age groups × 34 years × 33 boroughs.
    Returns tidy: gender, age, year, area, population."""
    if not _exists("MYEs_LA_1991_2024_gender.xlsx"):
        return pd.DataFrame(columns=["gender", "age", "year", "area", "population"])
    raw = pd.read_excel(_dp("MYEs_LA_1991_2024_gender.xlsx"), sheet_name="Data", header=None)
    age_map = {"Age 0 - 4": "0-4", "Aged 5-9": "5-9", "Aged 10-14": "10-14", "Aged 15-19": "15-19"}
    recs, i, n = [], 0, len(raw)
    while i < n:
        if str(raw.iloc[i, 0]) == "gender":
            gender = str(raw.iloc[i, 1]).strip()
            age = age_map.get(str(raw.iloc[i + 1, 1]).strip(), str(raw.iloc[i + 1, 1]).strip())
            areas = raw.iloc[i + 3].tolist()          # Date header row
            j = i + 4
            while j < n and pd.notna(raw.iloc[j, 0]) and str(raw.iloc[j, 0]).replace(".0", "").isdigit():
                year = int(float(raw.iloc[j, 0]))
                for k in range(1, len(areas)):
                    area = areas[k]
                    if pd.isna(area):
                        continue
                    val = pd.to_numeric(raw.iloc[j, k], errors="coerce")
                    recs.append((gender, age, year, str(area).strip(), val))
                j += 1
            i = j
        else:
            i += 1
    df = pd.DataFrame(recs, columns=["gender", "age", "year", "area", "population"])
    df["area"] = (df["area"].str.replace("and Fulham", "& Fulham", regex=False)
                            .str.replace("and Chelsea", "& Chelsea", regex=False))
    return df

@st.cache_data(show_spinner=False)
def load_mye_lsoa():
    """ONS small-area MYEs, Westminster LSOAs, mid-2022 → mid-2024 (already
    Westminster-only). Single-year F0..F90 / M0..M90 collapsed into the four
    child age bands × {Female, Male, Total}. Returns tidy long frame."""
    fn = "Small_Area_Output_Area_Mid_Year_Estimated.xlsx"
    if not _exists(fn):
        return pd.DataFrame()
    sheets = {"Mid-2022 LSOA 2021": 2022, "Mid-2023 LSOA 2021": 2023, "Mid-2024 LSOA 2021": 2024}
    bands = {"0-4": range(0, 5), "5-9": range(5, 10), "10-14": range(10, 15), "15-19": range(15, 20)}
    frames = []
    for sheet, year in sheets.items():
        try:
            df = pd.read_excel(_dp(fn), sheet_name=sheet, header=3)
        except Exception:
            continue
        df = df[df["LAD 2023 Name"] == "Westminster"].copy()
        base = df[["LSOA 2021 Code", "LSOA 2021 Name"]].rename(
            columns={"LSOA 2021 Code": "LSOA_CODE", "LSOA 2021 Name": "LSOA_NAME"})
        for band, rng in bands.items():
            f = df[[f"F{a}" for a in rng if f"F{a}" in df.columns]].sum(axis=1)
            m = df[[f"M{a}" for a in rng if f"M{a}" in df.columns]].sum(axis=1)
            for gender, vals in [("Female", f), ("Male", m), ("Total", f + m)]:
                rec = base.copy()
                rec["year"], rec["age"], rec["gender"] = year, band, gender
                rec["count"] = vals.values
                frames.append(rec)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_imd():
    """IoD 2025 IMD composite. National file filtered to Westminster LSOAs."""
    if not _exists("File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx"):
        return pd.DataFrame()
    df = pd.read_excel(_dp("File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx"), sheet_name="IMD25")
    rank = [c for c in df.columns if "IMD) Rank" in c][0]
    dec = [c for c in df.columns if "IMD) Decile" in c][0]
    df = df.rename(columns={"LSOA code (2021)": "LSOA_CODE", "LSOA name (2021)": "LSOA_NAME",
                            "Local Authority District name (2024)": "LAD", rank: "IMD_Rank", dec: "IMD_Decile"})
    return df[df["LAD"] == "Westminster"][["LSOA_CODE", "LSOA_NAME", "IMD_Rank", "IMD_Decile"]].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_idaci():
    """IoD 2025 IDACI (children) supplementary index, filtered to Westminster."""
    if not _exists("File_3_IoD2025_Supplementary_Indices_IDACI_and_IDAOPI.xlsx"):
        return pd.DataFrame()
    df = pd.read_excel(_dp("File_3_IoD2025_Supplementary_Indices_IDACI_and_IDAOPI.xlsx"),
                       sheet_name="IoD2025 IDACI & IDAOPI")
    rank = [c for c in df.columns if "IDACI) Rank" in c][0]
    dec = [c for c in df.columns if "IDACI) Decile" in c][0]
    df = df.rename(columns={"LSOA code (2021)": "LSOA_CODE", "LSOA name (2021)": "LSOA_NAME",
                            "Local Authority District name (2024)": "LAD",
                            rank: "IDACI_Rank", dec: "IDACI_Decile"})
    return df[df["LAD"] == "Westminster"][["LSOA_CODE", "LSOA_NAME", "IDACI_Rank", "IDACI_Decile"]].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_ward_geojson():
    """Westminster ward boundaries from TopoJSON (already WGS84), id = WardCode."""
    if not _exists("Wards_WCC.json"):
        return None
    with open(_dp("Wards_WCC.json")) as f:
        raw = json.load(f)
    gj = _topojson_to_geojson(raw) if raw.get("type") == "Topology" else raw
    gj = _ensure_wgs84(gj)
    for ft in gj["features"]:
        p = ft["properties"]
        p["WardName"] = p.get("WardName") or p.get("Name") or p.get("DisplayNam")
    return _inject_id(gj, "WardCode")

@st.cache_data(show_spinner=False)
def load_ward_lookup():
    """LSOA→ward lookup with coverage %. Used to label LSOAs with their ward and
    to build coverage-weighted ward averages. Empty (graceful) if file absent.
    Expected layout: column of E01… codes, then ward name, then coverage %."""
    if not _exists("Ward LSOA Lookup.xlsx"):
        return pd.DataFrame(columns=["LSOA_CODE", "Ward", "coverage"])
    raw = pd.read_excel(_dp("Ward LSOA Lookup.xlsx"), header=None)
    code_col = None
    for c in raw.columns:
        if raw[c].astype(str).str.match(r"E01\d+").any():
            code_col = c
            break
    if code_col is None:
        return pd.DataFrame(columns=["LSOA_CODE", "Ward", "coverage"])
    out = pd.DataFrame({
        "LSOA_CODE": raw[code_col].astype(str).str.strip(),
        "Ward": raw[code_col + 1].astype(str).str.strip()})
    cov = raw[code_col + 2].astype(str).str.replace("%", "", regex=False).str.strip()
    out["coverage"] = pd.to_numeric(cov, errors="coerce")
    out = out[out["LSOA_CODE"].str.match(r"E01\d+")].dropna(subset=["coverage"])
    if not out.empty and out["coverage"].max() <= 1.5:     # fraction → percentage
        out["coverage"] = out["coverage"] * 100
    return out.reset_index(drop=True)

def ward_for_lsoa(lookup):
    """LSOA→dominant ward (highest coverage) dict, for labelling/hover."""
    if lookup.empty:
        return {}
    return lookup.sort_values("coverage").groupby("LSOA_CODE")["Ward"].last().to_dict()

def coverage_weighted_ward(lookup, lsoa_df, value_col, code_col="LSOA_CODE"):
    """Aggregate an LSOA value to ward level via coverage weights:
    score_ward = Σ(coverage_i × value_i) / Σ(coverage_i) over the ward's LSOAs.
    e.g. E01004703 is 0.01% of Abbey Road → contributes 0.01% of the weight."""
    if lookup.empty or lsoa_df.empty:
        return pd.DataFrame(columns=["Ward", value_col])
    m = lookup.merge(lsoa_df[[code_col, value_col]], on=code_col, how="inner").dropna(subset=[value_col])
    if m.empty:
        return pd.DataFrame(columns=["Ward", value_col])
    g = m.groupby("Ward").apply(
        lambda d: np.average(d[value_col], weights=d["coverage"]) if d["coverage"].sum() else np.nan)
    return g.reset_index(name=value_col)

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADERS — existing datasets (resolved from data/ folder; degrade if absent)
# ══════════════════════════════════════════════════════════════════════════════
def _pct(s):
    v = pd.to_numeric(s.astype(str).str.replace("%", "", regex=False)
                       .str.replace(",", "", regex=False).str.strip(), errors="coerce")
    if v.dropna().size and v.dropna().max() <= 1.0:
        v = v * 100
    return v

def _num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(), errors="coerce")

@st.cache_data(show_spinner=False)
def load_low_income_la():
    """DWP CiLIF — Table 2: children (0–15) in relative low income (AHC), by LA, FYE24–25."""
    if not _exists("2_AHC_Relative_LA.csv"):
        return pd.DataFrame(columns=["LA", "Area_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"])
    df = pd.read_csv(_dp("2_AHC_Relative_LA.csv"), header=8,
                     names=["LA", "Area_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"],
                     usecols=[0, 1, 2, 3, 4, 5])
    df = df.dropna(subset=["Area_Code"])
    for c in ["Pct_2024", "Pct_2025"]:
        df[c] = _pct(df[c])
    for c in ["N_2024", "N_2025"]:
        df[c] = _num(df[c])
    return df

@st.cache_data(show_spinner=False)
def load_low_income_ward():
    """DWP CiLIF — Table 4: children (0–15) in relative low income (AHC), by ward, FYE24–25."""
    if not _exists("4_AHC_Relative_Ward.csv"):
        return pd.DataFrame(columns=["LA", "LA_Code", "Ward", "Ward_Code", "N_2024", "N_2025",
                                     "Pct_2024", "Pct_2025", "LA_filled"])
    df = pd.read_csv(_dp("4_AHC_Relative_Ward.csv"), header=9,
                     names=["LA", "LA_Code", "Ward", "Ward_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"],
                     usecols=[0, 1, 2, 3, 4, 5, 6, 7])
    df = df.dropna(subset=["Ward_Code"])
    for c in ["Pct_2024", "Pct_2025"]:
        df[c] = _pct(df[c])
    for c in ["N_2024", "N_2025"]:
        df[c] = _num(df[c])
    df["LA_filled"] = df["LA"].ffill()
    return df

@st.cache_data(show_spinner=False)
def load_ks4_ethnic():
    """KS4 Attainment 8 by ethnic group, inner-London LAs, 2024/25."""
    path = _dp("data-key-stage-4-performance.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["ethnic_group", "subgroup", "la", "att8_2425", "pct_5above"])
    df = pd.read_csv(path, header=None, low_memory=False)
    r2 = df.iloc[2, 4:].astype(str).ffill(); r3 = df.iloc[3, 4:].astype(str).ffill(); r4 = df.iloc[4, 4:].astype(str).ffill()
    cols, seen = ["ethnic_group", "subgroup", "region", "la"], {}
    for j in range(len(r2)):
        base = f"{r2.iloc[j].strip().replace('/', '_')}_{r3.iloc[j].strip().replace(' ', '_')[:22]}_{r4.iloc[j].strip().replace(' ', '_')[:18]}"
        seen[base] = seen.get(base, 0) + 1
        cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    data = df.iloc[5:].copy(); data.columns = cols[:len(data.columns)]
    for c in ["ethnic_group", "subgroup", "region"]:
        data[c] = data[c].ffill()
    inner = ["Camden", "Hackney", "Hammersmith and Fulham", "Haringey", "Islington",
             "Kensington and Chelsea", "Lambeth", "Lewisham", "Newham", "Southwark",
             "Tower Hamlets", "Wandsworth", "Westminster"]
    sub = data[data["la"].isin(inner) & data["subgroup"].astype(str).str.startswith("All")].copy()
    att_col = [c for c in sub.columns if "2024_25" in c and "Attainment_8" in c and "Total" in c]
    pct_col = [c for c in sub.columns if "2024_25" in c and "achieving_gr" in c and "Total" in c]
    sub["att8_2425"] = pd.to_numeric(sub[att_col[0]].astype(str).str.replace("no data", "", regex=False), errors="coerce") if att_col else np.nan
    sub["pct_5above"] = pd.to_numeric(sub[pct_col[0]].astype(str).str.replace("no data", "", regex=False).str.replace("%", "", regex=False), errors="coerce") if pct_col else np.nan
    sub["la"] = (sub["la"].astype(str).str.replace("and Fulham", "& Fulham", regex=False).str.replace("and Chelsea", "& Chelsea", regex=False))
    sub["ethnic_group"] = (sub["ethnic_group"].astype(str)
                           .str.replace("Asian / Asian British", "Asian", regex=False)
                           .str.replace("Black / African / Caribbean / Black British", "Black", regex=False)
                           .str.replace("Mixed / multiple ethnic groups", "Mixed", regex=False)
                           .str.replace("Other ethnic group", "Other", regex=False))
    return sub

@st.cache_data(show_spinner=False)
def load_ks4_time():
    """KS4 Attainment 8 all-pupils time series 2018/19–2024/25, inner-London LAs."""
    path = _dp("data-key-stage-4-performance__1_.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["la", "year", "att8"])
    df = pd.read_csv(path, header=None, low_memory=False)
    YEARS = ["2018/19", "2019/20", "2020/21", "2021/22", "2022/23", "2023/24", "2024/25"]
    ATT8_COLS = [12, 13, 14, 15, 16, 17, 18]
    for ci in [0, 1, 2, 3]:
        df.iloc[5:, ci] = df.iloc[5:, ci].ffill()
    inner = {"Camden", "Hackney", "Hammersmith and Fulham", "Haringey", "Islington",
             "Kensington and Chelsea", "Lambeth", "Lewisham", "Newham", "Southwark",
             "Tower Hamlets", "Wandsworth", "Westminster"}
    sub = df.iloc[5:].copy()
    mask = (sub.iloc[:, 4].isin(inner) & sub.iloc[:, 0].isna() & sub.iloc[:, 1].isna() & sub.iloc[:, 2].isna())
    recs = []
    for _, row in sub[mask].iterrows():
        la = str(row.iloc[4]).strip()
        for yr, cidx in zip(YEARS, ATT8_COLS):
            recs.append({"la": la, "year": yr,
                         "att8": pd.to_numeric(str(row.iloc[cidx]).replace(",", "").strip(), errors="coerce")})
    ts = pd.DataFrame(recs)
    if not ts.empty:
        ts["la"] = (ts["la"].str.replace("and Fulham", "& Fulham", regex=False).str.replace("and Chelsea", "& Chelsea", regex=False))
    return ts

# ── CENSUS 2021 (tidy long loaders; flexible to the Nomis export shape) ───────
def _lsoa_code(s):
    return s.astype(str).str.extract(r"(E\d{8,})")[0]

def _lsoa_name(s):
    return s.astype(str).str.replace(r"E\d{8,}\s*:?\s*", "", regex=True).str.strip()

@st.cache_data(show_spinner=False)
def load_rm006():
    """RM006 — Age of youngest dependent child by household type (Census 2021).
    Returns LONG: LSOA_CODE, LSOA_NAME, household_type, youngest_age, count.
    Handles a 2-row (household type / age) Nomis header; falls back to a flat
    age-only export tagged household_type='All households'."""
    fn = "RM006_age_of_youngest_dependent_child_by_household_type.xlsx"
    if not _exists(fn):
        return pd.DataFrame(columns=["LSOA_CODE", "LSOA_NAME", "household_type", "youngest_age", "count"])
    HH = ["One-person household", "Married or civil partnership couple household",
          "Cohabiting couple household", "Lone parent household", "Multi-person household"]
    AGES = {"No dependent children": "No dependent children",
            "Aged 0 to 4": "0 to 4", "Aged 5 to 9": "5 to 9",
            "Aged 10 to 15": "10 to 15", "Aged 16 to 18": "16 to 18"}
    try:
        head = pd.read_excel(_dp(fn), header=None, nrows=12)
    except Exception:
        return pd.DataFrame(columns=["LSOA_CODE", "LSOA_NAME", "household_type", "youngest_age", "count"])
    # locate the row that contains the LSOA-labelled data (first 'E0190…' style code)
    hdr_row = None
    for r in range(11):
        if head.iloc[r].astype(str).str.contains("youngest|Aged|dependent", case=False, na=False).any():
            hdr_row = r
    if hdr_row is None:
        hdr_row = 7
    raw = pd.read_excel(_dp(fn), header=None)
    # two header rows: household type (ffill) above age band
    hh_row = raw.iloc[hdr_row - 1].ffill().astype(str)
    age_row = raw.iloc[hdr_row].astype(str)
    data = raw.iloc[hdr_row + 1:].copy()
    first = data.iloc[:, 0].astype(str)
    data = data[first.str.contains(r"E\d{8,}", na=False)]
    recs = []
    for ci in range(1, raw.shape[1]):
        hh = next((h for h in HH if h.lower() in str(hh_row.iloc[ci]).lower()), None)
        age = next((a for k, a in AGES.items() if k.lower() in str(age_row.iloc[ci]).lower()), None)
        if age is None:
            continue
        hh = hh or "All households"
        recs.append((ci, hh, age))
    if not recs:                       # flat fallback: cols = 5 age bands, all households
        cols = ["No dependent children", "0 to 4", "5 to 9", "10 to 15", "16 to 18"]
        out = []
        for _, row in data.iterrows():
            code = pd.Series([str(row.iloc[0])]).pipe(_lsoa_code).iloc[0]
            name = pd.Series([str(row.iloc[0])]).pipe(_lsoa_name).iloc[0]
            for k, age in enumerate(cols, start=1):
                if k < raw.shape[1]:
                    out.append((code, name, "All households", age,
                                pd.to_numeric(row.iloc[k], errors="coerce")))
        return pd.DataFrame(out, columns=["LSOA_CODE", "LSOA_NAME", "household_type", "youngest_age", "count"]).dropna(subset=["LSOA_CODE"])
    out = []
    for _, row in data.iterrows():
        code = pd.Series([str(row.iloc[0])]).pipe(_lsoa_code).iloc[0]
        name = pd.Series([str(row.iloc[0])]).pipe(_lsoa_name).iloc[0]
        for ci, hh, age in recs:
            out.append((code, name, hh, age, pd.to_numeric(row.iloc[ci], errors="coerce")))
    return pd.DataFrame(out, columns=["LSOA_CODE", "LSOA_NAME", "household_type", "youngest_age", "count"]).dropna(subset=["LSOA_CODE"])

@st.cache_data(show_spinner=False)
def load_rm012():
    """RM012 — Dependent children by ethnic group of Household Reference Person (HRP)
    by age (Census 2021). Returns LONG: LSOA_CODE, LSOA_NAME, hrp_group, age_band, count.
    HRP groups: Asian / Black / Mixed / White / Other. Age bands: 0-2, 3-4, 5-11, 12-15, 16-18."""
    fn = "RM012_dependent_children_by_HRP_ethnic_group_by_age.xlsx"
    if not _exists(fn):
        # tolerant of alternative file names
        for alt in ["RM012.xlsx", "RM012_ethnic_group_HRP_age.xlsx"]:
            if _exists(alt):
                fn = alt
                break
        else:
            return pd.DataFrame(columns=["LSOA_CODE", "LSOA_NAME", "hrp_group", "age_band", "count"])
    HRP = {"Asian": "Asian", "Black": "Black", "Mixed": "Mixed", "White": "White", "Other": "Other"}
    AGES = {"0 to 2": "0-2", "3 to 4": "3-4", "5 to 11": "5-11", "12 to 15": "12-15", "16 to 18": "16-18"}
    raw = pd.read_excel(_dp(fn), header=None)
    hdr_row = None
    for r in range(12):
        if raw.iloc[r].astype(str).str.contains("Aged|to ", na=False).any():
            hdr_row = r
    if hdr_row is None:
        return pd.DataFrame(columns=["LSOA_CODE", "LSOA_NAME", "hrp_group", "age_band", "count"])
    hrp_row = raw.iloc[hdr_row - 1].ffill().astype(str)
    age_row = raw.iloc[hdr_row].astype(str)
    data = raw.iloc[hdr_row + 1:].copy()
    data = data[data.iloc[:, 0].astype(str).str.contains(r"E\d{8,}", na=False)]
    spec = []
    for ci in range(1, raw.shape[1]):
        grp = next((g for k, g in HRP.items() if k.lower() in str(hrp_row.iloc[ci]).lower()), None)
        age = next((a for k, a in AGES.items() if k.lower() in str(age_row.iloc[ci]).lower()), None)
        if grp and age:
            spec.append((ci, grp, age))
    out = []
    for _, row in data.iterrows():
        code = pd.Series([str(row.iloc[0])]).pipe(_lsoa_code).iloc[0]
        name = pd.Series([str(row.iloc[0])]).pipe(_lsoa_name).iloc[0]
        for ci, grp, age in spec:
            out.append((code, name, grp, age, pd.to_numeric(row.iloc[ci], errors="coerce")))
    return pd.DataFrame(out, columns=["LSOA_CODE", "LSOA_NAME", "hrp_group", "age_band", "count"]).dropna(subset=["LSOA_CODE"])

@st.cache_data(show_spinner=False)
def load_rm033():
    """RM033 — Ethnic group of the DEPENDENT CHILD by sex (Census 2021).
    Returns LONG: LSOA_CODE, LSOA_NAME, sex, eth_detail, eth_high, count.
    eth_detail keeps Nomis sub-categories (e.g. 'Bangladeshi'); eth_high rolls them
    up to White / Asian / Black / Mixed / Other / Arab."""
    fn = "RM033_ethic_group_dependent_child_by_sex.xlsx"
    if not _exists(fn):
        for alt in ["RM033_ethnic_group_dependent_child_by_sex.xlsx", "RM033.xlsx"]:
            if _exists(alt):
                fn = alt
                break
        else:
            return pd.DataFrame(columns=["LSOA_CODE", "LSOA_NAME", "sex", "eth_detail", "eth_high", "count"])
    raw = pd.read_excel(_dp(fn), header=None)
    # find header row carrying ethnicity labels (contains a colon hierarchy)
    hdr_row = None
    for r in range(12):
        if raw.iloc[r].astype(str).str.contains(":", na=False).sum() >= 3:
            hdr_row = r
            break
    if hdr_row is None:
        hdr_row = 8
    sex_row = raw.iloc[hdr_row - 1].ffill().astype(str) if hdr_row >= 1 else pd.Series([""] * raw.shape[1])
    eth_row = raw.iloc[hdr_row].astype(str)
    data = raw.iloc[hdr_row + 1:].copy()
    data = data[data.iloc[:, 0].astype(str).str.contains(r"E\d{8,}", na=False)]

    def high_level(label):
        l = label.lower()
        if "white" in l: return "White"
        if "asian" in l: return "Asian"
        if "black" in l: return "Black"
        if "mixed" in l or "multiple" in l: return "Mixed"
        if "arab" in l: return "Arab"
        return "Other"

    def sex_of(label):
        l = str(label).lower()
        if "female" in l: return "Female"
        if "male" in l: return "Male"
        return "All"

    spec = []
    for ci in range(1, raw.shape[1]):
        eth = str(eth_row.iloc[ci]).strip()
        if not eth or eth.lower() in ("nan", "total"):
            continue
        detail = eth.split(":")[-1].strip() or eth
        spec.append((ci, sex_of(sex_row.iloc[ci]), detail, high_level(eth)))
    out = []
    for _, row in data.iterrows():
        code = pd.Series([str(row.iloc[0])]).pipe(_lsoa_code).iloc[0]
        name = pd.Series([str(row.iloc[0])]).pipe(_lsoa_name).iloc[0]
        for ci, sex, detail, high in spec:
            out.append((code, name, sex, detail, high, pd.to_numeric(row.iloc[ci], errors="coerce")))
    return pd.DataFrame(out, columns=["LSOA_CODE", "LSOA_NAME", "sex", "eth_detail", "eth_high", "count"]).dropna(subset=["LSOA_CODE"])

# ── EGDI ──────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_egdi_lsoa():
    if not _exists("EGDI.xlsx"):
        return pd.DataFrame()
    df = pd.read_excel(_dp("EGDI.xlsx"), sheet_name="Data")
    edi = [c for c in df.columns if c.startswith("EDI.")]
    keep = ["LSOA21CD", "LSOA21NM", "Range", "Mostdeprivedgroup", "Leastdeprivedgroup",
            "TopEGDIDEC", "BottomEGDIDEC"] + edi
    return df[[c for c in keep if c in df.columns]].copy()

@st.cache_data(show_spinner=False)
def load_egdi():
    if not _exists("EGDI-Local-Authority-profiles.xlsx"):
        return pd.DataFrame()
    df = pd.read_excel(_dp("EGDI-Local-Authority-profiles.xlsx"), sheet_name="Profiles")
    df.columns = ["idx", "LA_Code", "LA_Name", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10",
                  "Total_LSOAs", "Pct_D1", "Pct_D2", "Pct_D3", "Pct_D4", "Pct_D5", "Pct_D6", "Pct_D7",
                  "Pct_D8", "Pct_D9", "Pct_D10", "_a", "_b", "_c", "Category", "_d", "Flat",
                  "More_ethnic_ineq", "Less_ethnic_ineq", "N_shape", "Pct_bottom20", "Pct_top20"][:df.shape[1]]
    return df.iloc[1:].reset_index(drop=True)

# ── GEOJSON loaders ───────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_lsoa_geojson():
    """Westminster 2021 LSOA boundaries → WGS84, id = LSOA21CD."""
    if not _exists("ONS_LSOA_2021 (1).json"):
        return None
    with open(_dp("ONS_LSOA_2021 (1).json")) as f:
        raw = json.load(f)
    gj = _topojson_to_geojson(raw) if raw.get("type") == "Topology" else raw
    gj = _ensure_wgs84(gj)
    id_prop = "LSOA21CD" if "LSOA21CD" in gj["features"][0]["properties"] else \
        next((k for k in gj["features"][0]["properties"] if "LSOA" in k and "CD" in k.upper()), "LSOA21CD")
    return _inject_id(gj, id_prop)

@st.cache_data(show_spinner=False)
def load_borough_geojson():
    """London borough boundaries → WGS84, id = borough code."""
    if not _exists("Borough_London_LL84.json"):
        return None
    with open(_dp("Borough_London_LL84.json")) as f:
        raw = json.load(f)
    gj = _topojson_to_geojson(raw) if raw.get("type") == "Topology" else raw
    gj = _ensure_wgs84(gj)
    props = gj["features"][0]["properties"]
    id_prop = next((k for k in props if "code" in k.lower() or k == "BoroughCod"), list(props)[0])
    return _inject_id(gj, id_prop)

# ══════════════════════════════════════════════════════════════════════════════
# LOAD EVERYTHING
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("Loading datasets…"):
    df_mye_la    = load_mye_la()
    df_mye_lsoa  = load_mye_lsoa()
    df_imd       = load_imd()
    df_idaci     = load_idaci()
    ward_gj      = load_ward_geojson()
    ward_lookup  = load_ward_lookup()
    lsoa_to_ward = ward_for_lsoa(ward_lookup)
    df_li_la     = load_low_income_la()
    df_li_ward   = load_low_income_ward()
    df_ks4_eth   = load_ks4_ethnic()
    df_ks4_ts    = load_ks4_time()
    df_rm006     = load_rm006()
    df_rm012     = load_rm012()
    df_rm033     = load_rm033()
    df_egdi      = load_egdi()
    df_egdi_lsoa = load_egdi_lsoa()
    lsoa_gj      = load_lsoa_geojson()
    borough_gj   = load_borough_geojson()

def add_ward(df, code_col="LSOA_CODE"):
    """Attach the LSOA's (dominant) ward as a 'Ward' column + a labelled name."""
    out = df.copy()
    out["Ward"] = out[code_col].map(lsoa_to_ward).fillna("—")
    if "LSOA_NAME" in out.columns:
        out["LSOA_labelled"] = np.where(out["Ward"] != "—",
                                        out["LSOA_NAME"] + " · " + out["Ward"], out["LSOA_NAME"])
    return out

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo = _dp("city_of_westminster.png")
    if os.path.exists(logo):
        st.image(logo, width=160)
    st.markdown("## Westminster Children's Dashboard")
    st.markdown("**CIPFA Statistical Neighbours**")
    for b in NEIGHBOURS:
        dot = FOCAL if b == "Westminster" else CONTEXT_BAR
        st.markdown(f"<span style='color:{dot}'>■</span> {b}", unsafe_allow_html=True)
    st.divider()
    st.caption("**How to read the charts**")
    st.markdown("- **Westminster is always in strong colour**; comparators are muted.\n"
                "- **Click a legend entry** to hide a line/bar; **double-click** to isolate one.\n"
                "- Every chart has a **PNG download** for slides.")
    st.divider()
    st.caption("**Data sources**")
    st.markdown("[ONS Mid-Year Population Estimates](https://www.ons.gov.uk/peoplepopulationandcommunity/populationandmigration/populationestimates)")
    st.markdown("[IoD 2025 — IMD & supplementary indices](https://www.gov.uk/government/statistics/english-indices-of-deprivation-2025)")
    st.markdown("[DWP Children in Low Income Families](https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025)")
    st.markdown("[DfE Key Stage 4 — Explore Education Statistics](https://explore-education-statistics.service.gov.uk/)")
    st.markdown("[EGDI — gedi.ac.uk](https://gedi.ac.uk/egdi/)")
    st.markdown("[Census 2021 — ONS Nomis](https://www.nomisweb.co.uk/)")

# ── HEADER + METRICS ──────────────────────────────────────────────────────────
st.title("🏙️ Westminster Children's Demographics")
st.markdown("Population, child poverty, attainment, ethnicity and deprivation — "
            "benchmarked against CIPFA statistical neighbours, with Westminster always in focus.")

# child population now (MYE LSOA, validated) — most up-to-date figure
child_now = np.nan
if not df_mye_lsoa.empty:
    child_now = int(df_mye_lsoa[(df_mye_lsoa["year"] == df_mye_lsoa["year"].max()) &
                                (df_mye_lsoa["gender"] == "Total")]["count"].sum())
# IDACI worst-10% share (validated)
idaci_share = round((df_idaci["IDACI_Decile"] == 1).mean() * 100, 1) if not df_idaci.empty else np.nan
imd_share = round((df_imd["IMD_Decile"] == 1).mean() * 100, 1) if not df_imd.empty else np.nan
# child poverty
wcc_li = df_li_la[df_li_la["LA"].astype(str).str.contains("Westminster", na=False)]
wcc_li = wcc_li.iloc[0] if len(wcc_li) else pd.Series({"N_2025": np.nan, "Pct_2025": np.nan,
                                                       "N_2024": np.nan, "Pct_2024": np.nan})

c1, c2, c3, c4 = st.columns(4)
c1.metric("Children aged 0–19 (MYE, mid-2024)",
          f"{child_now:,}" if not np.isnan(child_now) else "—",
          help="Latest ONS mid-year estimate, Westminster, ages 0–19. The most up-to-date child count.")
c2.metric("Children in low income (FYE 2025)",
          f"{int(wcc_li['N_2025']):,}" if pd.notna(wcc_li["N_2025"]) else "—",
          delta=(f"{wcc_li['Pct_2025'] - wcc_li['Pct_2024']:+.1f}pp vs FYE24" if pd.notna(wcc_li["Pct_2024"]) else None),
          delta_color="inverse",
          help="DWP AHC relative low income, children 0–15. Green ↓ = improvement.")
c3.metric("LSOAs in worst 10% — child income (IDACI 2025)",
          f"{idaci_share}%" if not np.isnan(idaci_share) else "—",
          help="Share of Westminster's 123 LSOAs in the most deprived national IDACI decile. "
               "Has almost doubled from ~11% (2019) to ~21% (2025).")
c4.metric("LSOAs in worst 10% — overall (IMD 2025)",
          f"{imd_share}%" if not np.isnan(imd_share) else "—",
          help="Share of Westminster LSOAs in the most deprived national IMD decile. "
               "IDACI (child-specific) reveals far more deprivation than the overall IMD.")
st.divider()

# ── TABS ──────────────────────────────────────────────────────────────────────
tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "🏠 Overview",
    "📍 Child Poverty",
    "🗺️ Population & Demographics",
    "📚 KS4 Attainment",
    "⚖️ Deprivation (IMD · IDACI · EGDI)",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — OVERVIEW / LANDING
# ══════════════════════════════════════════════════════════════════════════════
with tab0:
    st.subheader("What this dashboard shows — and which dataset to trust for what")
    st.markdown(
        "Westminster's child population is measured by several official sources, each built "
        "differently and each best for a particular question. This page explains what to use when, "
        "so the numbers later in the dashboard are read in the right context.")

    st.markdown("##### The datasets at a glance")
    st.markdown(f"""
<div class="ds-card"><b>① Mid-Year Estimates (MYEs) — start here for ages.</b><br>
ONS rebuilds the population every year from the Census, births, deaths and migration. They are the
<b>most up-to-date</b> count of children by age (0–4, 5–9, 10–14, 15–19) and sex, available right down
to LSOA level (mid-2022 → mid-2024) and back to 1991 at borough level. <b>Use MYEs as the default for
"how many children, what age, where, and how is that changing".</b> Their limitation: they carry
<b>no ethnicity and no household detail</b>.</div>

<div class="ds-card"><b>② Census 2021 — for the detail MYEs can't give.</b><br>
A full count once a decade. It is now a few years old, but it is the <b>only</b> source with the
granularity of <b>ethnicity</b> and <b>household type</b>. We include three Census tables:
<i>RM033</i> (the child's own ethnicity — the best ethnicity source), <i>RM012</i> (children by the
ethnicity of the household reference person, by age) and <i>RM006</i> (age of the youngest child by
household type). Use the Census when the question is specifically about ethnicity or household
structure; otherwise prefer the MYEs.</div>

<div class="ds-card"><b>③ Index of Multiple Deprivation (IMD 2025) + IDACI — the deprivation backbone.</b><br>
The IMD is the standard <b>overall</b> measure of neighbourhood deprivation. Its child-specific
supplementary index, <b>IDACI</b> (Income Deprivation Affecting Children), is the <b>first go-to for
child poverty geography</b> — it measures the share of children in income-deprived families and, on the
2025 after-housing-costs basis, exposes hidden child poverty in high-cost areas like Westminster.
<b>Prefer the IMD alongside IDACI.</b></div>

<div class="ds-card"><b>④ Ethnic Group Deprivation Index (EGDI) — the ethnicity dimension of deprivation.</b><br>
IMD and IDACI tell you <i>where</i> deprivation sits; they cannot tell you whether it falls
<b>unevenly across ethnic groups</b> within the same neighbourhood. The EGDI adds exactly that lens.
Reading it alongside IDACI and IMD turns "this area is deprived" into "and deprivation here is borne
disproportionately by particular ethnic groups" — essential for targeting support equitably.</div>
""", unsafe_allow_html=True)

    st.info("**In one line:** use the **Mid-Year Estimates** for the current age picture, the "
            "**Census** for ethnicity and household detail, the **IMD with IDACI** for deprivation "
            "(IDACI first for children), and the **EGDI** to see how deprivation is distributed "
            "across ethnic groups.")

    st.markdown("##### How to use the charts")
    st.markdown(
        "- **Colour carries the story.** Westminster is always in strong colour; comparator boroughs "
        "are deliberately muted so your eye lands on Westminster first.\n"
        "- **The legend is interactive.** Click a borough's name to hide it; double-click to show only "
        "that one. Drop Camden or Hammersmith & Fulham to declutter a busy line chart.\n"
        "- **Maps are interactive.** Hover any area for its name, its ward, and the value.\n"
        "- **Every chart exports as a PNG** (the ⬇ button, or the 📷 icon on the chart toolbar) so it "
        "drops straight into a slide deck.")
    legend_hint()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — CHILD POVERTY (DWP Children in Low Income Families)
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Children in low income families — Westminster vs CIPFA neighbours")
    st.markdown(
        "**Dataset:** DWP *Children in Low Income Families: local area statistics, 2022–2025*. "
        "The figures here use the **After-Housing-Costs (AHC) relative** measure for children aged "
        "**0–15** — the more meaningful measure in London because it accounts for very high housing "
        "costs. Two tables are used: **Table 2** (Local Authority) for the borough comparison and "
        "**Table 4** (Ward) for the within-Westminster breakdown. "
        "[Source & definitions](https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025).")

    if df_li_la.empty:
        st.info("Child-poverty LA file (`2_AHC_Relative_LA.csv`) not found in the data folder.")
    else:
        nb_codes = list(NEIGHBOURS.values())
        df_nb = df_li_la[df_li_la["Area_Code"].isin(nb_codes)].copy()
        df_nb["Borough"] = (df_nb["LA"].str.replace("and Fulham", "& Fulham", regex=False)
                                       .str.replace("and Chelsea", "& Chelsea", regex=False))
        df_nb = df_nb.sort_values("Pct_2025")

        col_a, col_b = st.columns([3, 2])
        with col_a:
            chart_title("Westminster sits mid-table among its neighbours for child poverty",
                        "% of children (0–15) in relative low income, after housing costs · FYE 2025")
            cmap = borough_palette(df_nb["Borough"])
            fig1 = px.bar(df_nb, x="Pct_2025", y="Borough", orientation="h",
                          color="Borough", color_discrete_map=cmap, text="Pct_2025")
            fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
            fig1.update_xaxes(range=[0, df_nb["Pct_2025"].max() * 1.22],
                              title="% children in low income (AHC)")
            fig1.update_yaxes(title="")
            show_chart(fig1, "child_poverty_bar", "DWP CiLIF, Table 2 (LA), FYE 2025")

        with col_b:
            chart_title("All neighbours improved 2024 → 2025",
                        "Westminster (in colour) against muted comparators")
            df_ts = df_nb[["Borough", "Pct_2024", "Pct_2025"]].melt("Borough", var_name="Year", value_name="Pct")
            df_ts["Year"] = df_ts["Year"].map({"Pct_2024": "FYE 2024", "Pct_2025": "FYE 2025"})
            fig2 = px.line(df_ts, x="Year", y="Pct", color="Borough", markers=True,
                           color_discrete_map=borough_palette(df_ts["Borough"].unique()))
            for tr in fig2.data:
                if tr.name == "Westminster":
                    tr.line.width = 4; tr.marker.size = 11
                else:
                    tr.line.width = 2; tr.line.color = CONTEXT_LINE; tr.marker.color = CONTEXT_LINE
            fig2.update_yaxes(title="% children in low income", rangemode="tozero")
            fig2.update_xaxes(title="")
            show_chart(fig2, "child_poverty_trend", "DWP CiLIF, Table 2 (LA), FYE 2024–25")
        legend_hint()

        # ── CIPFA choropleth
        st.divider()
        st.subheader("Geographic context — CIPFA statistical neighbours")
        st.markdown(
            "CIPFA *statistical neighbours* are the local authorities most similar to Westminster on "
            "socio-economic characteristics, so they are the fairest comparators. The map shades each "
            "neighbour by its child-poverty rate (AHC relative, FYE 2025); Westminster is outlined in "
            "the centre. Reading Westminster against this group — rather than against England as a whole "
            "— is the basis for the benchmarking throughout this dashboard.")
        if borough_gj is not None:
            chart_title("Child poverty across Westminster's CIPFA neighbours",
                        "% of children (0–15) in relative low income (AHC) · FYE 2025")
            nb_gj = {"type": "FeatureCollection",
                     "features": [f for f in borough_gj["features"] if f["id"] in nb_codes]}
            fig_m = choropleth(nb_gj, df_nb["Area_Code"].tolist(), df_nb["Pct_2025"].tolist(),
                               df_nb["Borough"].tolist(), "% in low income",
                               [[0, WCC["light_blue"]], [0.5, "#5B79C9"], [1.0, FOCAL]],
                               fmt=":.1f", zoom=10.3, center={"lat": 51.505, "lon": -0.17}, height=460)
            fig_m.data[0].hovertemplate = "<b>%{text}</b><br>% in low income: %{z:.1f}%<extra></extra>"
            show_chart(fig_m, "cipfa_map", "DWP CiLIF, Table 2 (LA), FYE 2025")
        else:
            st.info("Borough boundary file (`Borough_London_LL84.json`) not found — borough map skipped.")

    # ── Ward level
    st.divider()
    st.subheader("Within Westminster — ward-level child poverty (FYE 2025)")
    st.markdown(
        "**Dataset:** DWP *Children in Low Income Families*, **Table 4 (Ward)** — children aged 0–15, "
        "AHC relative measure. Child poverty is highly concentrated: the north-west of the borough "
        "carries far higher rates than the centre and south.")
    if df_li_ward.empty:
        st.info("Child-poverty ward file (`4_AHC_Relative_Ward.csv`) not found in the data folder.")
    else:
        wcc_w = df_li_ward[df_li_ward["LA_filled"].astype(str).str.contains("Westminster", na=False)].copy()
        wcc_w = wcc_w.dropna(subset=["Ward", "Pct_2025"]).sort_values("Pct_2025")
        top3 = wcc_w.nlargest(3, "Pct_2025")["Ward"].tolist()
        chart_title(f"Child poverty is concentrated in {', '.join(top3)}",
                    "% of children (0–15) in relative low income (AHC) · FYE 2025 · top three wards in colour")
        wcc_w["col"] = np.where(wcc_w["Ward"].isin(top3), FOCAL, CONTEXT_BAR)
        fig3 = go.Figure(go.Bar(
            x=wcc_w["Pct_2025"], y=wcc_w["Ward"], orientation="h",
            marker_color=wcc_w["col"], text=[f"{v:.1f}%" for v in wcc_w["Pct_2025"]],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>% children in low income: %{x:.1f}%<extra></extra>"))
        fig3.update_xaxes(range=[0, wcc_w["Pct_2025"].max() * 1.2], title="% children in low income (AHC)")
        fig3.update_yaxes(title="")
        show_chart(fig3, "ward_poverty", "DWP CiLIF, Table 4 (Ward), FYE 2025")
        st.success(f"💡 **Recommendation:** {top3[0]}, {top3[1]} and {top3[2]} should be the priority "
                   "wards for child-poverty interventions — the same north-west cluster that the IDACI "
                   "and IMD maps (Deprivation tab) independently flag as most deprived.")

    source_line("Child-poverty figures use the AHC (after-housing-costs) relative measure, children "
                "aged 0–15. FYE = financial year ending. Table 2 = Local Authority; Table 4 = Ward. "
                "Source: DWP/HMRC Children in Low Income Families local area statistics, 2022–2025.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — POPULATION & DEMOGRAPHICS
# Order (per guidance): MYE LSOA → MYE LA (borough + 1991–2024) → RM033 → RM012 → RM006
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Westminster LSOA demographic maps (Census 2021 & Mid-Year Estimates)")
    st.markdown(
        "This section maps **how many children live where, by age and sex**, then adds the "
        "**ethnicity and household detail** only the Census can provide. Start with the **Mid-Year "
        "Estimates** — they are the most up-to-date count — then use the Census tables for ethnicity "
        "and household structure.")
    legend_hint("On maps, use the filters to switch age band, sex or category; hover any area for its "
                "ward and value.")

    # ──────────────────────────────────────────────────────────────────────────
    # 2.1  MYE LSOA — the go-to current age picture
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 1 · Mid-Year Estimates — children by LSOA (mid-2022 → mid-2024)")
    st.markdown(
        "**Dataset:** ONS *small-area mid-year population estimates*, Westminster LSOAs. The MYEs are "
        "rebuilt every year, so this is the **most current** view of where children live. They cover "
        "age bands **0–4, 5–9, 10–14, 15–19** and sex, but carry **no ethnicity or household detail** "
        "(use the Census tables below for those).")
    if df_mye_lsoa.empty:
        st.info("Small-area MYE file not found — upload `Small_Area_Output_Area_Mid_Year_Estimated.xlsx`.")
    elif lsoa_gj is None:
        st.info("LSOA boundary file `ONS_LSOA_2021 (1).json` not found — the map needs it to draw areas.")
    else:
        cma, cmb, cmc = st.columns(3)
        yr = cma.selectbox("Year (mid-year)", sorted(df_mye_lsoa["year"].unique(), reverse=True),
                           key="mye_lsoa_yr")
        sx = cmb.selectbox("Sex", ["Total", "Female", "Male"], key="mye_lsoa_sex")
        ag = cmc.selectbox("Age band", ["All 0–19"] + MYE_AGE_ORDER, key="mye_lsoa_age")
        d = df_mye_lsoa[(df_mye_lsoa["year"] == yr) & (df_mye_lsoa["gender"] == sx)].copy()
        if ag != "All 0–19":
            d = d[d["age"] == ag]
        d = d.groupby(["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        d = add_ward(d)
        chart_title(f"Where Westminster's children live — {ag.lower() if ag!='All 0–19' else 'ages 0–19'}, {sx.lower()}, {yr}",
                    "ONS small-area mid-year estimates · darker = more children")
        fig = choropleth(lsoa_gj, d["LSOA_CODE"], d["count"], d["LSOA_NAME"],
                         "Children", [[0, WCC["light_blue"]], [1, FOCAL]],
                         wards=d["Ward"].tolist(), fmt=":,")
        show_chart(fig, "mye_lsoa_map", "ONS small-area MYEs, Westminster LSOAs")
        st.caption(f"Total {('ages 0–19' if ag=='All 0–19' else ag)} ({sx.lower()}), {yr}: "
                   f"**{int(d['count'].sum()):,}** children across {d['LSOA_CODE'].nunique()} LSOAs.")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 2.2  MYE LA — London-borough choropleth + 1991–2024 trend
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 2 · Mid-Year Estimates — London boroughs & long-run trend (1991–2024)")
    st.markdown(
        "**Dataset:** ONS *mid-year population estimates*, all 33 London boroughs, **1991–2024**. This "
        "puts Westminster's child numbers in a London-wide context and shows how they have changed over "
        "three decades. Filter by sex and age band; the map shows the chosen year, the line chart the "
        "full back-series.")
    if df_mye_la.empty:
        st.info("Borough MYE file not found — upload `MYEs_LA_1991_2024_gender.xlsx`.")
    else:
        f1, f2, f3 = st.columns(3)
        g_la = f1.selectbox("Sex", ["Total", "Female", "Male"], key="mye_la_sex")
        a_la = f2.selectbox("Age band", ["All 0–19"] + MYE_AGE_ORDER, key="mye_la_age")
        y_la = f3.selectbox("Year (map)", sorted(df_mye_la["year"].unique(), reverse=True), key="mye_la_yr")

        dla = df_mye_la[df_mye_la["gender"] == g_la].copy()
        if a_la != "All 0–19":
            dla = dla[dla["age"] == a_la]

        # London-borough choropleth (selected year)
        dmap = dla[dla["year"] == y_la].groupby("area", as_index=False)["population"].sum()
        if borough_gj is not None and not dmap.empty:
            # map borough names → geojson ids
            id_by_name = {}
            for ft in borough_gj["features"]:
                p = ft["properties"]
                nm = (p.get("name") or p.get("BoroughNa") or p.get("NAME") or
                      next((v for k, v in p.items() if isinstance(v, str) and "E09" not in v), ""))
                id_by_name[str(nm).replace("and Fulham", "& Fulham").replace("and Chelsea", "& Chelsea")] = ft["id"]
            dmap["gid"] = dmap["area"].map(id_by_name)
            dmap2 = dmap.dropna(subset=["gid"])
            chart_title(f"London boroughs — children {('0–19' if a_la=='All 0–19' else a_la)}, {g_la.lower()}, {y_la}",
                        "ONS mid-year estimates · Westminster outlined")
            figm = choropleth(borough_gj, dmap2["gid"], dmap2["population"], dmap2["area"],
                              "Children", [[0, "#EDEFF6"], [1, FOCAL]], fmt=":,",
                              zoom=9, center={"lat": 51.50, "lon": -0.12}, height=520)
            show_chart(figm, "mye_la_map", "ONS mid-year population estimates, London boroughs")
        else:
            # fallback: ranked bar with Westminster focal
            chart_title(f"London boroughs — children {('0–19' if a_la=='All 0–19' else a_la)}, {g_la.lower()}, {y_la}",
                        "ONS mid-year estimates · Westminster in colour")
            dmap = dmap.sort_values("population", ascending=True)
            dmap["col"] = np.where(dmap["area"].str.contains("Westminster"), FOCAL, CONTEXT_BAR)
            figb = go.Figure(go.Bar(x=dmap["population"], y=dmap["area"], orientation="h",
                                    marker_color=dmap["col"],
                                    hovertemplate="<b>%{y}</b><br>Children: %{x:,}<extra></extra>"))
            figb.update_layout(height=640)
            figb.update_yaxes(title="")
            show_chart(figb, "mye_la_bar", "ONS mid-year population estimates (borough map file absent)")

        # 1991–2024 trend, CIPFA neighbours, Westminster focal
        chart_title(f"Three decades of change — children {('0–19' if a_la=='All 0–19' else a_la)}, {g_la.lower()}, 1991–2024",
                    "ONS mid-year estimates · Westminster in strong colour, CIPFA neighbours muted")
        figt = go.Figure()
        for b in NEIGHBOURS:
            sb = dla[dla["area"].str.contains(b.replace("& ", "").split()[0], case=False, na=False)]
            sb = dla[dla["area"] == b] if (dla["area"] == b).any() else sb
            ser = sb.groupby("year", as_index=False)["population"].sum().sort_values("year")
            if ser.empty:
                continue
            focal = (b == "Westminster")
            figt.add_trace(go.Scatter(
                x=ser["year"], y=ser["population"], mode="lines", name=b,
                line=dict(color=FOCAL if focal else CONTEXT_LINE, width=3.5 if focal else 1.5),
                opacity=1.0 if focal else 0.9,
                hovertemplate="<b>"+b+"</b><br>%{x}: %{y:,} children<extra></extra>"))
        figt.update_xaxes(title="Year", dtick=5)
        figt.update_yaxes(title="Children")
        show_chart(figt, "mye_la_trend", "ONS mid-year population estimates, 1991–2024")
        legend_hint("Drop comparator boroughs from the legend to isolate Westminster's trajectory.")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 2.3  RM033 — child's OWN ethnicity (best ethnicity source)
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 3 · Census RM033 — ethnicity of the child (the best ethnicity source)")
    st.markdown(
        "**Dataset:** Census 2021 *RM033 — ethnic group of the dependent child by sex*. This is the "
        "**most policy-relevant** ethnicity table because it records the **child's own ethnicity**, not "
        "a proxy. Use the high-level summary map first, then drill into a specific detailed group "
        "(e.g. *Bangladeshi*, *African*) and, if needed, split by sex.")
    if df_rm033.empty:
        st.info("RM033 file not found — upload the Census `RM033 … dependent child by sex` table to "
                "render these maps. (The loader and maps are ready; they just need the file.)")
    elif lsoa_gj is None:
        st.info("LSOA boundary file not found — needed to draw the ethnicity maps.")
    else:
        # Summary map: high-level groups, both sexes summed
        st.markdown("**Summary — high-level ethnic groups (both sexes)**")
        s1, s2 = st.columns(2)
        hi = s1.selectbox("Ethnic group (high level)", sorted(df_rm033["eth_high"].unique()), key="rm033_hi")
        d = df_rm033[(df_rm033["eth_high"] == hi)]
        d = d[d["sex"].isin(["All"])] if (d["sex"] == "All").any() else d
        d = d.groupby(["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        d = add_ward(d)
        chart_title(f"Children identifying as {hi} — Census 2021",
                    "RM033 (child's own ethnicity) · darker = more children · hover for ward")
        fig = choropleth(lsoa_gj, d["LSOA_CODE"], d["count"], d["LSOA_NAME"],
                         f"{hi} children", [[0, WCC["light_blue"]], [1, FOCAL]],
                         wards=d["Ward"].tolist(), fmt=":,")
        show_chart(fig, "rm033_summary", "Census 2021 RM033")

        # Detailed map: sub-category + sex
        st.markdown("**Detailed — specific ethnic sub-group, optionally by sex**")
        c1, c2 = st.columns(2)
        det = c1.selectbox("Detailed ethnic group", sorted(df_rm033["eth_detail"].unique()), key="rm033_det")
        sexes = sorted(df_rm033["sex"].unique())
        sx = c2.selectbox("Sex", sexes, index=sexes.index("All") if "All" in sexes else 0, key="rm033_sex")
        dd = df_rm033[(df_rm033["eth_detail"] == det) & (df_rm033["sex"] == sx)]
        dd = dd.groupby(["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        dd = add_ward(dd)
        chart_title(f"{det} children ({sx.lower()}) — Census 2021",
                    "RM033 detailed sub-category · darker = more children")
        figd = choropleth(lsoa_gj, dd["LSOA_CODE"], dd["count"], dd["LSOA_NAME"],
                          f"{det}", [[0, WCC["light_blue"]], [1, FOCAL_ALT]],
                          wards=dd["Ward"].tolist(), fmt=":,")
        show_chart(figd, "rm033_detail", "Census 2021 RM033")
        source_line("RM033 records the dependent child's own ethnic group — preferred over household-based "
                    "ethnicity (RM012 below) when the question is about the children themselves.")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 2.4  RM012 — children by ethnicity of HRP, by age
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 4 · Census RM012 — children by household ethnicity & age")
    st.markdown(
        "**Dataset:** Census 2021 *RM012 — dependent children by the ethnic group of the Household "
        "Reference Person (HRP), by age*. The **HRP** is the person in whose name the home is owned or "
        "rented (or the higher earner) — so this table groups children by their **household's** "
        "ethnicity rather than their own. It is most useful when the question is about household "
        "context and age structure together. Age bands: **0–2, 3–4, 5–11, 12–15, 16–18**.")
    if df_rm012.empty:
        st.info("RM012 file not found — upload the Census `RM012 … by HRP ethnic group by age` table. "
                "(Loader and maps are ready.)")
    elif lsoa_gj is None:
        st.info("LSOA boundary file not found — needed to draw these maps.")
    else:
        # Summary: all ages 0–18, by HRP group (summary categories first)
        st.markdown("**Summary — all children 0–18, by household (HRP) ethnic group**")
        hg = st.selectbox("HRP ethnic group", sorted(df_rm012["hrp_group"].unique()), key="rm012_hi")
        d = df_rm012[df_rm012["hrp_group"] == hg].groupby(
            ["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        d = add_ward(d)
        chart_title(f"Children in {hg}-HRP households (all ages 0–18) — Census 2021",
                    "RM012 · darker = more children · hover for ward")
        fig = choropleth(lsoa_gj, d["LSOA_CODE"], d["count"], d["LSOA_NAME"],
                         "Children", [[0, WCC["light_blue"]], [1, FOCAL]],
                         wards=d["Ward"].tolist(), fmt=":,")
        show_chart(fig, "rm012_summary", "Census 2021 RM012")

        # Detailed: HRP group + age band
        st.markdown("**Detailed — household ethnic group × age band**")
        c1, c2 = st.columns(2)
        hg2 = c1.selectbox("HRP ethnic group ", sorted(df_rm012["hrp_group"].unique()), key="rm012_g2")
        ab = c2.selectbox("Age band", sorted(df_rm012["age_band"].unique()), key="rm012_age")
        dd = df_rm012[(df_rm012["hrp_group"] == hg2) & (df_rm012["age_band"] == ab)]
        dd = dd.groupby(["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        dd = add_ward(dd)
        chart_title(f"{hg2}-HRP households, children aged {ab} — Census 2021",
                    "RM012 detailed · darker = more children")
        figd = choropleth(lsoa_gj, dd["LSOA_CODE"], dd["count"], dd["LSOA_NAME"],
                          "Children", [[0, WCC["light_blue"]], [1, FOCAL_ALT]],
                          wards=dd["Ward"].tolist(), fmt=":,")
        show_chart(figd, "rm012_detail", "Census 2021 RM012")
        source_line("HRP = Household Reference Person. RM012 classifies children by the ethnicity of "
                    "their household's reference person, so it differs from RM033 (the child's own ethnicity).")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 2.5  RM006 — age of youngest child by household type (LEAST important; last)
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 5 · Census RM006 — household type & age of the youngest child")
    st.markdown(
        "**Dataset:** Census 2021 *RM006 — age of the youngest dependent child by household type*. This "
        "is the **least central** table for child demographics, so it sits last. It is useful for "
        "understanding **family structure** — for example where lone-parent households with very young "
        "children are concentrated. Household types: one-person, married/civil-partnership couple, "
        "cohabiting couple, lone parent, and multi-person.")
    if df_rm006.empty:
        st.info("RM006 file not found — upload the Census `RM006 … youngest dependent child by household "
                "type` table. (Loader and maps are ready.)")
    elif lsoa_gj is None:
        st.info("LSOA boundary file not found — needed to draw these maps.")
    else:
        st.markdown("**Summary — all household types, by age of youngest child**")
        ya = st.selectbox("Age of youngest child", sorted(df_rm006["youngest_age"].unique()), key="rm006_age")
        d = df_rm006[df_rm006["youngest_age"] == ya].groupby(
            ["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        d = add_ward(d)
        chart_title(f"Households whose youngest child is {ya} — Census 2021",
                    "RM006, all household types · darker = more households")
        fig = choropleth(lsoa_gj, d["LSOA_CODE"], d["count"], d["LSOA_NAME"],
                         "Households", [[0, WCC["light_blue"]], [1, FOCAL]],
                         wards=d["Ward"].tolist(), fmt=":,")
        show_chart(fig, "rm006_summary", "Census 2021 RM006")

        st.markdown("**Detailed — household type × age of youngest child**")
        c1, c2 = st.columns(2)
        ht = c1.selectbox("Household type", sorted(df_rm006["household_type"].unique()), key="rm006_ht")
        ya2 = c2.selectbox("Age of youngest child ", sorted(df_rm006["youngest_age"].unique()), key="rm006_age2")
        dd = df_rm006[(df_rm006["household_type"] == ht) & (df_rm006["youngest_age"] == ya2)]
        dd = dd.groupby(["LSOA_CODE", "LSOA_NAME"], as_index=False)["count"].sum()
        dd = add_ward(dd)
        chart_title(f"{ht} — youngest child {ya2} — Census 2021",
                    "RM006 detailed · darker = more households")
        figd = choropleth(lsoa_gj, dd["LSOA_CODE"], dd["count"], dd["LSOA_NAME"],
                          "Households", [[0, WCC["light_blue"]], [1, FOCAL_ALT]],
                          wards=dd["Ward"].tolist(), fmt=":,")
        show_chart(figd, "rm006_detail", "Census 2021 RM006")
        source_line("RM006 counts households by the age of their youngest dependent child. Source: "
                    "ONS Census 2021, accessed via Nomis. Link to the table from the source list in the sidebar.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — KS4 ATTAINMENT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Key Stage 4 attainment — Westminster vs inner-London")
    st.markdown(
        "**Dataset:** DfE *Key Stage 4 performance*, accessed via Explore Education Statistics. "
        "**Attainment 8** is a pupil's average grade across eight core GCSE subjects (max 90). Here it "
        "is shown for Westminster against inner-London local authorities and its CIPFA neighbours, by "
        "ethnic group, over time, and across boroughs. Westminster is always in strong colour; "
        "comparators are muted.")
    legend_hint()

    # Attainment 8 by ethnic group (Westminster focal)
    if df_ks4_eth.empty:
        st.info("KS4 ethnicity file `data-key-stage-4-performance.csv` not found in the data folder.")
    else:
        wcc_e = df_ks4_eth[df_ks4_eth["la"].astype(str).str.contains("Westminster", na=False)].copy()
        wcc_e = wcc_e.dropna(subset=["att8_2425"]).sort_values("att8_2425")
        if not wcc_e.empty:
            chart_title("Attainment 8 in Westminster varies widely by ethnic group (2024/25)",
                        "Average Attainment 8 score, Westminster pupils · DfE KS4")
            wcc_e["col"] = FOCAL
            fige = go.Figure(go.Bar(
                x=wcc_e["att8_2425"], y=wcc_e["ethnic_group"], orientation="h",
                marker_color=wcc_e["col"], text=[f"{v:.1f}" for v in wcc_e["att8_2425"]],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Attainment 8: %{x:.1f}<extra></extra>"))
            fige.update_xaxes(title="Average Attainment 8 score")
            fige.update_yaxes(title="")
            show_chart(fige, "ks4_ethnic", "DfE KS4 performance, 2024/25")

    # CIPFA / inner-London bar (Westminster focal, others muted)
    if not df_ks4_eth.empty:
        allsub = df_ks4_eth[df_ks4_eth["ethnic_group"].astype(str).str.contains("Total|All", case=False, na=False)]
        la_att = (df_ks4_eth.dropna(subset=["att8_2425"])
                  .groupby("la", as_index=False)["att8_2425"].mean())
        if not la_att.empty:
            la_att = la_att.sort_values("att8_2425")
            la_att["col"] = np.where(la_att["la"].str.contains("Westminster"), FOCAL, CONTEXT_BAR)
            chart_title("Westminster against inner-London peers — Attainment 8 (2024/25)",
                        "Average Attainment 8 score · Westminster in colour, peers muted")
            figc = go.Figure(go.Bar(
                x=la_att["att8_2425"], y=la_att["la"], orientation="h",
                marker_color=la_att["col"], text=[f"{v:.1f}" for v in la_att["att8_2425"]],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Attainment 8: %{x:.1f}<extra></extra>"))
            figc.update_xaxes(title="Average Attainment 8 score")
            figc.update_yaxes(title="")
            figc.update_layout(height=460)
            show_chart(figc, "ks4_cipfa", "DfE KS4 performance, 2024/25")

            # Borough choropleth of Attainment 8
            if borough_gj is not None:
                id_by_name = {}
                for ft in borough_gj["features"]:
                    p = ft["properties"]
                    nm = (p.get("name") or p.get("BoroughNa") or p.get("NAME") or
                          next((v for k, v in p.items() if isinstance(v, str) and "E09" not in v), ""))
                    id_by_name[str(nm).replace("and Fulham", "& Fulham").replace("and Chelsea", "& Chelsea")] = ft["id"]
                la_att["gid"] = la_att["la"].map(id_by_name)
                geo = la_att.dropna(subset=["gid"])
                if not geo.empty:
                    chart_title("Attainment 8 across inner-London boroughs (2024/25)",
                                "DfE KS4 · darker = higher Attainment 8")
                    figm = choropleth(borough_gj, geo["gid"], geo["att8_2425"], geo["la"],
                                      "Attainment 8", [[0, "#EDEFF6"], [1, FOCAL]], fmt=":.1f",
                                      zoom=9.2, center={"lat": 51.51, "lon": -0.12}, height=520)
                    show_chart(figm, "ks4_map", "DfE KS4 performance, 2024/25")

    # Attainment 8 trend (remove yellow → Westminster focal, peers muted)
    if df_ks4_ts.empty:
        st.info("KS4 time-series file `data-key-stage-4-performance__1_.csv` not found.")
    else:
        chart_title("Attainment 8 over time — Westminster vs inner-London",
                    "Average Attainment 8 score · Westminster in strong colour, peers muted")
        figt = go.Figure()
        for la in sorted(df_ks4_ts["la"].unique()):
            s = df_ks4_ts[df_ks4_ts["la"] == la].dropna(subset=["att8"]).sort_values("year")
            if s.empty:
                continue
            focal = "Westminster" in la
            figt.add_trace(go.Scatter(
                x=s["year"], y=s["att8"], mode="lines+markers" if focal else "lines", name=la,
                line=dict(color=FOCAL if focal else CONTEXT_LINE, width=3.5 if focal else 1.3),
                opacity=1.0 if focal else 0.85,
                hovertemplate="<b>"+la+"</b><br>%{x}: %{y:.1f}<extra></extra>"))
        figt.update_xaxes(title="Academic year")
        figt.update_yaxes(title="Average Attainment 8 score")
        show_chart(figt, "ks4_trend", "DfE KS4 performance, 2018/19–2024/25")
        legend_hint("Hide peer boroughs from the legend to read Westminster's trend on its own.")
        source_line("Attainment 8 measures pupils' average achievement across eight GCSE subjects. "
                    "Source: DfE Key Stage 4 performance, Explore Education Statistics. See sidebar for link.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — DEPRIVATION (IMD · IDACI · EGDI)
# ══════════════════════════════════════════════════════════════════════════════
N_ENGLAND_LSOA = 33755   # England LSOAs (2021) — for national deprivation percentiles

with tab4:
    st.subheader("Deprivation — three complementary lenses")
    st.markdown(
        "Three official measures are used together here, because each answers a different question:\n\n"
        "- **IMD (Index of Multiple Deprivation 2025)** — the standard **overall** measure of "
        "neighbourhood deprivation, combining income, employment, education, health, crime, housing and "
        "environment into one composite rank. *Where is this neighbourhood deprived overall?*\n"
        "- **IDACI (Income Deprivation Affecting Children Index)** — a child-specific supplementary "
        "index: the **proportion of children aged 0–15 living in income-deprived families**. This is the "
        "**first go-to for child poverty geography**. *Where are children specifically affected?*\n"
        "- **EGDI (Ethnic Group Deprivation Index)** — shows whether deprivation within a neighbourhood "
        "falls **unevenly across ethnic groups**. *Is deprivation here shared, or borne disproportionately "
        "by particular ethnic groups?*\n\n"
        "Read IDACI **first** for children, the IMD alongside it for the overall picture, and the EGDI "
        "to see the ethnic distribution of that deprivation.")
    legend_hint("On every map, hover an area for its ward and value; darker shading = more deprived.")

    # ──────────────────────────────────────────────────────────────────────────
    # 4.1  IDACI — children in income-deprived families (first go-to)
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 1 · IDACI — income deprivation affecting children (first go-to)")
    st.markdown(
        "**Dataset:** IoD 2025 *Income Deprivation Affecting Children Index* — the share of children "
        "aged **0–15** in income-deprived families (benefit-unit basis). Lower national **rank** and "
        "lower **decile** mean **more** deprived (decile 1 = worst 10% in England).")

    st.markdown(
        "> **What the 2025 IDACI shows for Westminster.** Child income-deprivation has risen sharply. "
        "**Church Street** is now the **28th most deprived ward in England** for child poverty at **81%** "
        "(top 0.5% nationally), up **+47 percentage points since 2019**. **Westbourne (73%)**, "
        "**Harrow Road (63%)**, **Queen's Park (63%)** and **Pimlico South (62%)** all sit in the worst "
        "**6%** nationally. The **biggest increases** are **Pimlico South (+67ppt, to 99%)** and "
        "**Little Venice (+64ppt, to 98%)**. Six small areas now exceed **90%** child poverty — **half of "
        "them in Church Street**. The share of Westminster LSOAs in the **worst 10% nationally has almost "
        "doubled, from ~11% (2019) to ~21% (2025)**; the borough's average IDACI score rose from 18.3% to "
        "**43.5%**, moving its national rank from 96th to **57th of 296**. London boroughs now account for "
        "**5 of the 10** most child-deprived in England.")

    if df_idaci.empty:
        st.info("IDACI file not found — upload `File_3_IoD2025_Supplementary_Indices_IDACI_and_IDAOPI.xlsx`.")
    elif lsoa_gj is None:
        st.info("LSOA boundary file `ONS_LSOA_2021 (1).json` not found — needed to draw the IDACI map.")
    else:
        di = df_idaci.copy()
        # national deprivation percentile (higher = more deprived) from rank
        di["dep_pctile"] = (1 - (di["IDACI_Rank"] - 1) / N_ENGLAND_LSOA) * 100
        di = add_ward(di)
        # LSOA choropleth (decile; reverse so decile 1 = darkest)
        chart_title("Child income-deprivation across Westminster's LSOAs (IDACI 2025)",
                    "National IDACI decile · 1 = worst 10% in England · darker = more deprived · hover for ward")
        figi = choropleth(lsoa_gj, di["LSOA_CODE"], di["IDACI_Decile"], di["LSOA_NAME"],
                          "IDACI decile (1=worst)", [[0, WCC["light_blue"]], [1, FOCAL]],
                          wards=di["Ward"].tolist(), fmt=":.0f", reverse=True)
        show_chart(figi, "idaci_lsoa", "IoD 2025 IDACI, Westminster LSOAs")

        # Ward-level coverage-weighted map
        st.markdown("**Ward-level IDACI** (LSOA values aggregated to wards using the coverage-weighted "
                    "LSOA→ward lookup — e.g. an LSOA only 0.01% inside Abbey Road contributes 0.01% of the weight).")
        if ward_gj is None:
            st.info("Ward boundary file `Wards_WCC.json` not found.")
        elif ward_lookup.empty:
            st.info("Ward LSOA lookup (`Ward LSOA Lookup.xlsx`) not found — needed for coverage-weighted "
                    "ward aggregation. The LSOA map above is unaffected; the ward map will render once the "
                    "lookup is supplied.")
        else:
            wagg = coverage_weighted_ward(ward_lookup, di, "dep_pctile")
            # map ward names → geojson ward ids
            name_to_id = {}
            for ft in ward_gj["features"]:
                p = ft["properties"]
                name_to_id[str(p.get("WardName") or p.get("Name"))] = ft["id"]
            wagg["wid"] = wagg["Ward"].map(name_to_id)
            wg = wagg.dropna(subset=["wid"])
            if not wg.empty:
                chart_title("Child income-deprivation by ward (IDACI 2025, coverage-weighted)",
                            "Mean national deprivation percentile of the ward's LSOAs · darker = more deprived")
                figw = choropleth(ward_gj, wg["wid"], wg["dep_pctile"], wg["Ward"],
                                  "Deprivation percentile", [[0, WCC["light_blue"]], [1, FOCAL]],
                                  fmt=":.0f", zoom=11.5, height=520)
                show_chart(figw, "idaci_ward", "IoD 2025 IDACI aggregated to wards (coverage-weighted)")

        # Worst-10% decile distribution bar (decile 1 highlighted)
        dist = (di.groupby("IDACI_Decile").size().reindex(range(1, 11), fill_value=0)
                / len(di) * 100).reset_index()
        dist.columns = ["decile", "pct"]
        dist["col"] = np.where(dist["decile"] == 1, FOCAL, CONTEXT_BAR)
        chart_title(f"{(di['IDACI_Decile']==1).mean()*100:.0f}% of Westminster's LSOAs are in England's worst 10% for child income-deprivation",
                    "Distribution of Westminster's 123 LSOAs across national IDACI deciles (1 = most deprived)")
        figd = go.Figure(go.Bar(
            x=dist["decile"], y=dist["pct"], marker_color=dist["col"],
            text=[f"{v:.0f}%" for v in dist["pct"]], textposition="outside",
            hovertemplate="National decile %{x}<br>%{y:.1f}% of Westminster LSOAs<extra></extra>"))
        figd.update_xaxes(title="National IDACI decile (1 = most deprived 10%)", dtick=1)
        figd.update_yaxes(title="% of Westminster LSOAs")
        show_chart(figd, "idaci_dist", "IoD 2025 IDACI, Westminster LSOAs")
        source_line("IDACI = the proportion of children aged 0–15 in income-deprived families. "
                    "Ranks/deciles are national (England, 33,755 LSOAs). Published ward proportions above are "
                    "from the IoD 2025 release. Source: MHCLG Indices of Deprivation 2025 (see sidebar).")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 4.2  IMD — overall composite
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 2 · IMD — overall neighbourhood deprivation (the composite)")
    st.markdown(
        "**Dataset:** IoD 2025 *Index of Multiple Deprivation* — the standard **overall** measure. It "
        "combines **seven domains** (income, employment, education, health, crime, barriers to housing & "
        "services, and living environment) into a single national rank, where **1 = most deprived**. "
        "Because it averages across all of life, the IMD can **understate child-specific deprivation** in "
        "high-cost areas — which is exactly why IDACI (above) is the first go-to for children. Comparing "
        "the two maps shows where child poverty is worse than the overall picture suggests.")
    if df_imd.empty:
        st.info("IMD file not found — upload `File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx`.")
    elif lsoa_gj is None:
        st.info("LSOA boundary file not found — needed to draw the IMD map.")
    else:
        dm = add_ward(df_imd.copy())
        chart_title("Overall deprivation across Westminster's LSOAs (IMD 2025)",
                    "National IMD decile · 1 = worst 10% in England · darker = more deprived · hover for ward")
        figimd = choropleth(lsoa_gj, dm["LSOA_CODE"], dm["IMD_Decile"], dm["LSOA_NAME"],
                            "IMD decile (1=worst)", [[0, WCC["light_blue"]], [1, FOCAL]],
                            wards=dm["Ward"].tolist(), fmt=":.0f", reverse=True)
        show_chart(figimd, "imd_lsoa", "IoD 2025 IMD, Westminster LSOAs")
        imd_share_t = (dm["IMD_Decile"] == 1).mean() * 100
        idaci_share_t = (df_idaci["IDACI_Decile"] == 1).mean() * 100 if not df_idaci.empty else np.nan
        if pd.notna(idaci_share_t):
            st.info(f"**IMD vs IDACI:** **{imd_share_t:.0f}%** of Westminster's LSOAs fall in the worst "
                    f"national decile on the **overall IMD**, but **{idaci_share_t:.0f}%** do on the "
                    f"**child-specific IDACI** — confirming that child income-deprivation is **more** "
                    f"widespread than the headline deprivation measure implies.")
        source_line("IMD 2025 combines seven weighted domains into one composite rank. Source: MHCLG "
                    "English Indices of Deprivation 2025 (see sidebar for link).")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 4.3  EGDI — ethnic group deprivation
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 3 · EGDI — how deprivation is distributed across ethnic groups")
    st.markdown(
        "**Dataset:** *Ethnic Group Deprivation Index (EGDI)*. Where the IMD and IDACI tell you **how "
        "deprived** a neighbourhood is overall, the EGDI tells you whether that deprivation is **shared "
        "evenly across ethnic groups** or **concentrated** on particular ones living in the same area. "
        "It does this by scoring deprivation **separately for each ethnic group within an LSOA**, so two "
        "areas with an identical IMD rank can look very different through the EGDI lens. This adds the "
        "**ethnic-inequality dimension** that the IMD and IDACI alone cannot show.")

    # 4.3a  LA classification
    if df_egdi.empty:
        st.info("EGDI local-authority profile file `EGDI-Local-Authority-profiles.xlsx` not found — "
                "the classification cards and radar below will render once it is supplied.")
    else:
        wrow = df_egdi[df_egdi["LA_Name"].astype(str).str.contains("Westminster", na=False)]
        if not wrow.empty:
            w = wrow.iloc[0]
            cat = str(w.get("Category", "—"))
            flat = str(w.get("Flat", "—"))
            pflat = None
            for cand in ["Pct_bottom20", "Pct_top20"]:
                if cand in w.index and pd.notna(w.get(cand)):
                    pflat = w.get(cand)
            chart_title(f"Westminster's EGDI profile is classified as “{cat}”",
                        "EGDI local-authority classification · Westminster in colour")
            k1, k2, k3 = st.columns(3)
            k1.metric("EGDI classification", cat)
            k2.metric("Profile shape", flat if flat not in ("nan", "—") else "See chart")
            if pd.notna(w.get("Total_LSOAs", np.nan)):
                k3.metric("LSOAs assessed", f"{int(w['Total_LSOAs'])}")
            st.markdown(
                "**What the classification means.** The EGDI groups local authorities by the *shape* of "
                "their ethnic-deprivation distribution. A **“flat”** profile means deprivation is spread "
                "fairly **evenly** across ethnic groups — no single group carries a disproportionate share. "
                "A **steep** or **N-shaped** profile means certain ethnic groups are markedly more deprived "
                "than others in the same neighbourhoods. Reading the bar below shows, LSOA by LSOA, how wide "
                "that gap between the most- and least-deprived ethnic group actually is.")

            # Visualise "flat": decile distribution of LSOAs (proportion, not count)
            decs = [c for c in df_egdi.columns if c.startswith("Pct_D")]
            if decs:
                vals = pd.to_numeric(w[decs], errors="coerce").values
                chart_title("What “flat” looks like — Westminster LSOAs across EGDI deciles",
                            "% of Westminster LSOAs in each EGDI decile · an even spread ⇒ a flat profile")
                figf = go.Figure(go.Bar(
                    x=list(range(1, len(vals) + 1)), y=vals, marker_color=FOCAL,
                    text=[f"{v:.0f}%" if pd.notna(v) else "" for v in vals], textposition="outside",
                    hovertemplate="EGDI decile %{x}<br>%{y:.1f}% of LSOAs<extra></extra>"))
                figf.update_xaxes(title="EGDI decile", dtick=1)
                figf.update_yaxes(title="% of Westminster LSOAs")
                show_chart(figf, "egdi_flat", "EGDI local-authority profiles")

    # 4.3b  Range-by-LSOA bar (with ward detail)
    if df_egdi_lsoa.empty:
        st.info("EGDI LSOA file `EGDI.xlsx` not found — the LSOA range bar, maps and per-category maps "
                "below will render once it is supplied. (All loaders and charts are ready.)")
    else:
        el = df_egdi_lsoa.rename(columns={"LSOA21CD": "LSOA_CODE", "LSOA21NM": "LSOA_NAME"}).copy()
        el = add_ward(el)
        if "Range" in el.columns:
            er = el.dropna(subset=["Range"]).sort_values("Range", ascending=True).tail(25)
            chart_title("Where the ethnic-deprivation gap is widest (top 25 LSOAs)",
                        "EGDI range = gap between the most- and least-deprived ethnic group in the LSOA · ward shown on hover")
            lbl = np.where(er["Ward"] != "—", er["LSOA_NAME"] + " · " + er["Ward"], er["LSOA_NAME"])
            figr = go.Figure(go.Bar(
                x=er["Range"], y=lbl, orientation="h", marker_color=FOCAL,
                customdata=er["Ward"],
                hovertemplate="<b>%{y}</b><br>EGDI range: %{x:.2f}<extra></extra>"))
            figr.update_xaxes(title="EGDI range (within-LSOA gap across ethnic groups)")
            figr.update_yaxes(title="")
            figr.update_layout(height=640)
            show_chart(figr, "egdi_range", "EGDI, Westminster LSOAs")
            st.markdown("The widest gaps cluster in the same north-west wards that the IDACI and IMD maps "
                        "flag — but here the story is **inequality between ethnic groups within** those "
                        "neighbourhoods, not just their overall deprivation.")

        # 4.3c  EGDI LSOA map + ward map
        edi_cols = [c for c in el.columns if c.startswith("EDI.")]
        metric_opts = (["Range"] if "Range" in el.columns else []) + edi_cols
        if metric_opts and lsoa_gj is not None:
            st.markdown("**EGDI maps** — choose the overall within-LSOA range, or an individual ethnic "
                        "category's EGDI score.")
            msel = st.selectbox("Metric / ethnic category",
                                metric_opts,
                                format_func=lambda c: "Overall range (gap across groups)" if c == "Range"
                                else c.replace("EDI.", "").replace(".", " "),
                                key="egdi_metric")
            md = el.dropna(subset=[msel])
            nice = "ethnic-deprivation range" if msel == "Range" else msel.replace("EDI.", "").replace(".", " ") + " EGDI"
            chart_title(f"Westminster LSOAs — {nice} (EGDI)",
                        "Darker = greater ethnic-group deprivation · hover for ward")
            figm = choropleth(lsoa_gj, md["LSOA_CODE"], md[msel], md["LSOA_NAME"],
                              nice, [[0, WCC["light_blue"]], [1, FOCAL]],
                              wards=md["Ward"].tolist(), fmt=":.2f")
            show_chart(figm, "egdi_lsoa_map", "EGDI, Westminster LSOAs")

            # ward-level coverage-weighted EGDI map
            st.markdown("**Ward-level EGDI** (coverage-weighted from the LSOA→ward lookup).")
            if ward_gj is None:
                st.info("Ward boundary file not found.")
            elif ward_lookup.empty:
                st.info("Ward LSOA lookup not found — the coverage-weighted ward EGDI map will render once "
                        "`Ward LSOA Lookup.xlsx` is supplied.")
            else:
                wa = coverage_weighted_ward(ward_lookup, md, msel)
                name_to_id = {str(ft["properties"].get("WardName") or ft["properties"].get("Name")): ft["id"]
                              for ft in ward_gj["features"]}
                wa["wid"] = wa["Ward"].map(name_to_id)
                wg = wa.dropna(subset=["wid"])
                if not wg.empty:
                    chart_title(f"Ward-level {nice} (EGDI, coverage-weighted)",
                                "LSOA EGDI scores aggregated to wards by coverage share · darker = more deprived")
                    figww = choropleth(ward_gj, wg["wid"], wg[msel], wg["Ward"],
                                       nice, [[0, WCC["light_blue"]], [1, FOCAL]],
                                       fmt=":.2f", zoom=11.5, height=520)
                    show_chart(figww, "egdi_ward_map", "EGDI aggregated to wards (coverage-weighted)")

            st.caption("This per-category map replaces the old ethnic heatmap: pick any ethnic category "
                       "above to see *that group's* deprivation geography across Westminster, LSOA by LSOA.")
        source_line("EGDI scores deprivation separately for each ethnic group within an LSOA; the range is "
                    "the gap between the most- and least-deprived group. Source: Ethnic Group Deprivation "
                    "Index (gedi.ac.uk). Ward figures are coverage-weighted LSOA→ward aggregates.")

    st.divider()
    # ──────────────────────────────────────────────────────────────────────────
    # 4.4  CIPFA deprivation profile radar (PROPORTION of LSOAs per decile)
    # ──────────────────────────────────────────────────────────────────────────
    st.markdown("### 4 · CIPFA deprivation profile — Westminster vs its statistical neighbours")
    st.markdown(
        "**Dataset:** IoD 2025 IMD, by local authority. This radar shows, for each CIPFA neighbour, the "
        "**proportion (%) of its LSOAs** falling in each national IMD decile. Proportions are used rather "
        "than counts because boroughs differ in size — a percentage profile makes the **shape** of "
        "deprivation comparable regardless of how many LSOAs a borough has.")

    @st.cache_data(show_spinner=False)
    def _imd_decile_profile():
        """% of each CIPFA borough's LSOAs in each national IMD decile (from File_1)."""
        if not _exists("File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx"):
            return pd.DataFrame()
        df = pd.read_excel(_dp("File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx"), sheet_name="IMD25")
        dec = [c for c in df.columns if "IMD) Decile" in c][0]
        lad = "Local Authority District name (2024)"
        names = {b.replace("& ", "and "): b for b in NEIGHBOURS}     # match file spelling
        sub = df[df[lad].isin(list(names))].copy()
        sub["borough"] = sub[lad].map(names)
        g = sub.groupby(["borough", dec]).size().rename("n").reset_index()
        g.columns = ["borough", "decile", "n"]
        g["pct"] = g["n"] / g.groupby("borough")["n"].transform("sum") * 100
        return g[["borough", "decile", "pct"]]

    prof = _imd_decile_profile()
    if prof.empty:
        st.info("IMD file not found — the deprivation-profile radar needs "
                "`File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx`.")
    else:
        deciles = list(range(1, 11))
        wide = (prof.pivot_table(index="borough", columns="decile", values="pct", fill_value=0)
                .reindex(columns=deciles, fill_value=0))
        chart_title("Deprivation profile — share of each borough's LSOAs by IMD decile",
                    "% of LSOAs in each national decile (1 = most deprived) · Westminster in strong colour")
        figr = go.Figure()
        theta = [f"Decile {d}" for d in deciles] + ["Decile 1"]
        for b in wide.index:
            if b == "Westminster":
                continue
            r = wide.loc[b].tolist()
            figr.add_trace(go.Scatterpolar(r=r + [r[0]], theta=theta, name=b,
                                           line=dict(color=CONTEXT_LINE, width=1.4), opacity=0.8))
        if "Westminster" in wide.index:
            r = wide.loc["Westminster"].tolist()
            figr.add_trace(go.Scatterpolar(r=r + [r[0]], theta=theta, name="Westminster",
                                           line=dict(color=FOCAL, width=3.5), fill="toself",
                                           fillcolor="rgba(11,34,101,0.12)"))
        figr.update_layout(polar=dict(radialaxis=dict(ticksuffix="%", angle=90)), height=560)
        show_chart(figr, "imd_radar", "IoD 2025 IMD, CIPFA neighbours")
        legend_hint("Click neighbours off in the legend to compare Westminster with one borough at a time.")

        # most-similar neighbour (Euclidean distance on decile-% vectors)
        if "Westminster" in wide.index and len(wide) > 1:
            w = wide.loc["Westminster"].values.astype(float)
            dists = {b: float(np.sqrt(((wide.loc[b].values.astype(float) - w) ** 2).sum()))
                     for b in wide.index if b != "Westminster"}
            closest = min(dists, key=dists.get)
            st.success(f"💡 **Most similar profile:** of its CIPFA neighbours, **{closest}** has the "
                       f"deprivation shape closest to Westminster's (smallest difference across the decile "
                       f"distribution). The muted lines furthest from Westminster's are the least alike.")
        source_line("Profiles use the proportion (%) of each borough's LSOAs in each national IMD decile, "
                    "so size differences between boroughs don't distort the comparison. Source: MHCLG "
                    "Indices of Deprivation 2025.")

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("Westminster Children's Demographics · Westminster always shown in strong colour, comparators "
           "muted (Economist “grey-the-context” style) · every chart exports to PNG for slides · data "
           "sources linked in the sidebar.")
