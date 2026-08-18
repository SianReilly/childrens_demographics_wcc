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

import os, io, json, copy, glob, re, warnings
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
    # census tables — accept whatever RM0xx… variant is actually committed
    "RM012_dependent_children_by_ethnic_group_of_HRP.xlsx": [
        "RM012_dependent_children_by_ethnic_group_of_HRP.xlsx",
        "RM012_dependent_children_by_HRP_ethnic_group_by_age.xlsx", "RM012.xlsx"],
    "RM006_age_of_youngest_dependent_child_by_household_type.xlsx": [
        "RM006_age_of_youngest_dependent_child_by_household_type.xlsx", "RM006.xlsx"],
    "RM033_ethic_group_dependent_child_by_sex.xlsx": [
        "RM033_ethic_group_dependent_child_by_sex.xlsx",
        "RM033_ethnic_group_dependent_child_by_sex.xlsx", "RM033.xlsx"],
}

def _stem_key(s):
    """Normalise a filename for fuzzy matching: lowercase, strip non-alphanumerics."""
    return re.sub(r"[^a-z0-9]", "", os.path.splitext(str(s))[0].lower())

def _dp(name):
    aliases = list(dict.fromkeys(_ALIASES.get(name, []) + [name]))
    # 1) exact / alias match
    for d in _SEARCH_DIRS:
        for a in aliases:
            p = os.path.join(d, a)
            if os.path.exists(p):
                return p
    # 2) fuzzy fallback — match the leading token (e.g. 'RM012') or the full stem
    ext = os.path.splitext(name)[1].lower()
    want_keys = {_stem_key(a) for a in aliases}
    lead = re.match(r"[A-Za-z]+\d+", name)          # e.g. 'RM012'
    lead = lead.group(0).lower() if lead else None
    for d in _SEARCH_DIRS:
        if not os.path.isdir(d):
            continue
        for f in sorted(glob.glob(os.path.join(d, "*"))):
            b = os.path.basename(f)
            if ext and os.path.splitext(b)[1].lower() != ext:
                continue
            k = _stem_key(b)
            if k in want_keys or (lead and k.startswith(lead)):
                return f
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

# Fixed, consistent colour per CIPFA borough for ALL borough-comparator charts
# (bars, lines, radar). Westminster is the only strong colour; the five neighbours
# get distinct but deliberately muted/desaturated hues so they're each identifiable
# yet never compete with Westminster. Maps and many-bar charts keep the blue scheme.
BOROUGH_COLOURS = {
    "Westminster":          FOCAL,       # strong WCC navy
    "Kensington & Chelsea": "#5FA8A0",   # muted teal
    "Camden":               "#D0A44C",   # muted gold
    "Hammersmith & Fulham": "#C57B8A",   # muted rose
    "Islington":            "#8598CE",   # muted periwinkle
    "Wandsworth":           "#94B36A",   # muted olive
}
# spare muted hues if a comparator outside the CIPFA six ever appears
_EXTRA_MUTED = ["#B08E6A", "#7FB0A0", "#A97BA5", "#6E9BB3", "#C0906B", "#8AA98A"]

# Benchmark averages — deliberately neutral greys, drawn dashed, so they read as
# "context lines" rather than as another borough.
LONDON, ENGLAND = "London", "England"
AVERAGE_COLOURS = {LONDON: "#7A8390", ENGLAND: "#4A5058"}
COMPARATORS = list(NEIGHBOURS) + [LONDON, ENGLAND]

def _norm_la(name):
    """Normalise local-authority names to the app's house spelling ('&' not 'and')."""
    s = str(name).strip()
    s = s.replace("Kensington and Chelsea", "Kensington & Chelsea")
    s = s.replace("Hammersmith and Fulham", "Hammersmith & Fulham")
    return s

def is_average(name):
    return str(name) in AVERAGE_COLOURS

def line_style(name):
    """Westminster solid+thick, boroughs solid, London/England dashed."""
    if name == "Westminster":
        return dict(width=3.8)
    if is_average(name):
        return dict(width=2.2, dash="dash")
    return dict(width=1.8)

def borough_palette(categories, focal="Westminster"):
    """Distinct but muted colour per borough (consistent everywhere); Westminster strong.
    London/England benchmarks get neutral greys."""
    out, ei = {}, 0
    for c in categories:
        cc = _norm_la(c)
        if cc == focal:
            out[c] = FOCAL
        elif cc in AVERAGE_COLOURS:
            out[c] = AVERAGE_COLOURS[cc]
        elif cc in BOROUGH_COLOURS:
            out[c] = BOROUGH_COLOURS[cc]
        else:
            out[c] = _EXTRA_MUTED[ei % len(_EXTRA_MUTED)]; ei += 1
    return out

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="WCC Children's Demographics",
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
    m = m.copy()
    m["_wv"] = m[value_col] * m["coverage"]
    g = m.groupby("Ward", as_index=False).agg(_wv=("_wv", "sum"), _w=("coverage", "sum"))
    g[value_col] = np.where(g["_w"] > 0, g["_wv"] / g["_w"], np.nan)
    return g[["Ward", value_col]]

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

def _read_stacked_census(fn, block_key, age_map, cat_clean=None, max_hdr_scan=8):
    """Parse a Nomis 'stacked block' export. Each block starts on a row whose first
    cell contains `block_key` (e.g. 'household type'); that row's second cell is the
    category value (e.g. 'One-person household'). A header row inside the block holds
    the age-band labels across columns; the rows beneath are LSOA × counts. Blocks are
    stacked down the sheet (this is the shape RM006 uses — a household type every
    ~136 rows). Returns a list of (LSOA_CODE, LSOA_NAME, category, age, count) or None
    if the sheet is not in this stacked shape (so the caller can try a wide parser)."""
    raw = pd.read_excel(_dp(fn), header=None)
    n, ncol = raw.shape
    col0 = raw.iloc[:, 0].fillna("").astype(str).str.strip().str.lower()
    bk = block_key.lower()
    starts = [i for i in range(n) if col0.iloc[i].startswith(bk)]
    if not starts:
        return None
    starts_ext = starts + [n]
    recs = []
    for bi in range(len(starts)):
        r0, r1 = starts[bi], starts_ext[bi + 1]
        cat = str(raw.iloc[r0, 1]).strip() if ncol > 1 else ""
        if not cat or cat.lower() == "nan":
            continue
        if cat_clean:
            cat = cat_clean(cat)
        age_cols, age_row = {}, None
        for rr in range(r0 + 1, min(r0 + 1 + max_hdr_scan, r1)):
            row = raw.iloc[rr].fillna("").astype(str)
            hits = {}
            for ci in range(1, ncol):
                lab = row.iloc[ci].strip().lower()
                for key, val in age_map.items():
                    if key in lab:
                        hits[ci] = val
                        break
            if len(hits) >= 2:
                age_cols, age_row = hits, rr
                break
        if not age_cols:
            continue
        for rr in range(age_row + 1, r1):
            c0 = str(raw.iloc[rr, 0])
            c1 = str(raw.iloc[rr, 1]) if ncol > 1 else ""
            code = _lsoa_code(pd.Series([c0 + " " + c1])).iloc[0]
            nm0 = _lsoa_name(pd.Series([c0])).iloc[0]
            name = nm0 if (nm0 and nm0.lower() not in ("nan", "")) else (
                c1 if re.search(r"[A-Za-z]", c1) else "")
            key_code = code if (isinstance(code, str) and code) else name
            if not key_code or str(key_code).lower() == "nan":
                continue
            for ci, age in age_cols.items():
                recs.append((key_code, name, cat, age,
                             pd.to_numeric(raw.iloc[rr, ci], errors="coerce")))
    return recs

@st.cache_data(show_spinner=False)
def load_rm006():
    """RM006 — Age of youngest dependent child by household type (Census 2021).
    Returns LONG: LSOA_CODE, LSOA_NAME, household_type, youngest_age, count.
    Handles a 2-row (household type / age) Nomis header; falls back to a flat
    age-only export tagged household_type='All households'."""
    fn = "RM006_age_of_youngest_dependent_child_by_household_type.xlsx"
    COLS = ["LSOA_CODE", "LSOA_NAME", "household_type", "youngest_age", "count"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    # AGE labels (order longest/most-specific first to avoid partial mis-matches)
    AGE_MAP = {"no dependent": "No dependent children", "10 to 15": "10 to 15",
               "16 to 18": "16 to 18", "5 to 9": "5 to 9", "0 to 4": "0 to 4",
               "10-15": "10 to 15", "16-18": "16 to 18", "5-9": "5 to 9", "0-4": "0 to 4"}
    # PRIMARY: the real RM006 shape — a stacked block per household type
    try:
        recs = _read_stacked_census(fn, "household type", AGE_MAP)
    except Exception:
        recs = None
    if recs:
        df = pd.DataFrame(recs, columns=COLS).dropna(subset=["count"])
        if df["household_type"].nunique() > 1:
            return df.reset_index(drop=True)
    # FALLBACK: older wide (2-row header) / flat exports
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
    fn = "RM012_dependent_children_by_ethnic_group_of_HRP.xlsx"
    COLS = ["LSOA_CODE", "LSOA_NAME", "hrp_group", "age_band", "count"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    HRP = {"Asian": "Asian", "Black": "Black", "Mixed": "Mixed", "White": "White", "Other": "Other"}
    AGES = {"0 to 2": "0-2", "3 to 4": "3-4", "5 to 11": "5-11", "12 to 15": "12-15", "16 to 18": "16-18"}

    def _hrp_clean(label):
        l = str(label).lower()
        for k, v in HRP.items():
            if k.lower() in l:
                return v
        return str(label).strip()

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
    wide = pd.DataFrame(out, columns=COLS).dropna(subset=["LSOA_CODE"])
    if not wide.empty:
        return wide
    # FALLBACK: stacked-block layout (one block per HRP ethnic group)
    try:
        recs = _read_stacked_census(fn, "ethnic group", AGES, cat_clean=_hrp_clean)
    except Exception:
        recs = None
    if recs:
        return pd.DataFrame(recs, columns=COLS).dropna(subset=["count"]).reset_index(drop=True)
    return wide

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
# NEW DATASETS — children & schools, population change, childcare costs
# Sources: DfE Explore Education Statistics (pupils, SEN, childcare provider survey),
# ONS internal-migration detailed estimates, ONS/Nomis births.
# Every loader degrades to an empty frame so a missing file only hides its own charts.
# ══════════════════════════════════════════════════════════════════════════════
def _academic_year_label(tp):
    """201516 -> '2015/16'; 2025 -> '2025'."""
    s = str(int(tp))
    return f"{s[:4]}/{s[4:]}" if len(s) == 6 else s

def _year_start(tp):
    """201516 -> 2015 (sortable numeric start year)."""
    s = str(int(tp))
    return int(s[:4])

def _fetch_csv(url, timeout=25):
    """Fetch a CSV over HTTP with a hard timeout. Returns a DataFrame or None.
    Used for the Nomis API links; the app never blocks if the API is unreachable."""
    try:
        import urllib.request
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return pd.read_csv(io.BytesIO(r.read()))
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def load_pupils():
    """DfE 'School pupils and their characteristics' — headcount/FTE by LA and phase.
    File covers Westminster + its five neighbours only (no London/England rows);
    London and England totals come from the SEN file instead (see load_sen)."""
    fn = "data-school-pupils-and-their-characteristics.csv"
    COLS = ["year", "year_label", "la", "phase", "headcount", "fte", "full_time", "part_time"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    df = pd.read_csv(_dp(fn), low_memory=False)
    df = df[df["geographic_level"].astype(str).str.lower() == "local authority"].copy()
    out = pd.DataFrame({
        "year": df["time_period"].map(_year_start),
        "year_label": df["time_period"].map(_academic_year_label),
        "la": df["la_name"].map(_norm_la),
        "phase": df["phase_type_grouping"].astype(str).str.strip(),
        "headcount": _num(df["headcount"]),
        "fte": _num(df["fte"]),
        "full_time": _num(df["full_time"]),
        "part_time": _num(df["part_time"]),
    })
    return out.dropna(subset=["headcount"]).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_sen():
    """DfE 'Special educational needs in England'. IMPORTANT: within each phase,
    sen_provision == 'Total' is the TOTAL pupil headcount (the denominator), while
    'Education, health and care plan' and 'SEN support / SEN without an EHC plan'
    are the SEN counts. The file carries National (England), Regional (London) and
    all 33 London LAs, so it supplies the London/England comparators for pupil
    numbers as well as the SEN analysis."""
    fn = "SEN_data-special-educational-needs-in-england.csv"
    COLS = ["year", "year_label", "geo", "phase", "provision", "count"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    df = pd.read_csv(_dp(fn), low_memory=False)
    lvl = df["geographic_level"].astype(str).str.lower()
    geo = np.where(lvl == "national", ENGLAND,
          np.where(lvl == "regional", df["region_name"].astype(str), df["la_name"].astype(str)))
    out = pd.DataFrame({
        "year": df["time_period"].map(_year_start),
        "year_label": df["time_period"].map(_academic_year_label),
        "geo": pd.Series(geo).map(_norm_la),
        "phase": df["phase_type_grouping"].astype(str).str.strip(),
        "provision": df["sen_provision"].astype(str).str.strip(),
        "count": _num(df["pupil_count"]),
    })
    out = out[out["geo"].astype(str).str.lower() != "nan"]
    return out.dropna(subset=["count"]).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_cross_border():
    """DfE cross-border movement: where an LA's resident pupils actually go to school.
    resident_headcount = pupils living in the LA; school_in_la / school_outside_la split
    that by where they are educated; the trailing borough columns are the destination
    breakdown (how many of this LA's residents attend school in each named borough)."""
    fn = "cross_border_data_data-school-pupils-and-their-characteristics.csv"
    COLS = ["year", "year_label", "la", "phase", "resident", "in_la", "out_la"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS), pd.DataFrame(columns=["year", "la", "destination", "pupils"])
    df = pd.read_csv(_dp(fn), low_memory=False)
    phase_col = "phase-type_grouping" if "phase-type_grouping" in df.columns else "phase_type_grouping"
    base = pd.DataFrame({
        "year": df["time_period"].map(_year_start),
        "year_label": df["time_period"].map(_academic_year_label),
        "la": df["la_name"].map(_norm_la),
        "phase": df[phase_col].astype(str).str.strip(),
        "resident": _num(df["resident_headcount"]),
        "in_la": _num(df["school_in_la"]),
        "out_la": _num(df["school_outside_la"]),
    })
    dest_cols = [c for c in df.columns if c in
                 ["Camden", "Hammersmith_and_Fulham", "Islington", "Kensington_and_Chelsea",
                  "Wandsworth", "Westminster"]]
    flows = []
    for c in dest_cols:
        flows.append(pd.DataFrame({
            "year": df["time_period"].map(_year_start),
            "la": df["la_name"].map(_norm_la),
            "phase": df[phase_col].astype(str).str.strip(),
            "destination": _norm_la(c.replace("_", " ")),
            "pupils": _num(df[c]),
        }))
    fl = pd.concat(flows, ignore_index=True) if flows else pd.DataFrame(columns=["year", "la", "phase", "destination", "pupils"])
    return base.dropna(subset=["resident"]).reset_index(drop=True), fl.dropna(subset=["pupils"]).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_childcare_costs():
    """DfE Childcare and early years provider survey — hourly fees by London LA,
    for 2-year-olds and 3-and-4-year-olds."""
    fn = "costs_data-childcare-and-early-years-provider-survey.csv"
    COLS = ["year", "la", "la_code", "child_age", "mean_fee", "median_fee"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    df = pd.read_csv(_dp(fn), low_memory=False)
    out = pd.DataFrame({
        "year": _num(df["time_period"]).astype("Int64"),
        "la": df["la_name"].map(_norm_la),
        "la_code": df["new_la_code"].astype(str).str.strip(),
        "child_age": df["child_age"].astype(str).str.strip(),
        "mean_fee": _num(df["mean_hourly_fee"]),
        "median_fee": _num(df["median_hourly_fee"]),
    })
    return out.dropna(subset=["mean_fee", "median_fee"], how="all").reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_indep_calc():
    """Westminster's own independent-school estimate: LA resident headcount (ONS)
    minus maintained-school resident headcount (cross-border movement)."""
    fn = "Independent_schools_calculations.xlsx"
    COLS = ["year", "year_label", "la", "phase", "mtd_resident", "mtd_in_la",
            "mtd_out_la", "la_resident", "indep_estimate"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    raw = pd.read_excel(_dp(fn), header=None)
    hdr = None
    for r in range(min(10, len(raw))):
        if raw.iloc[r].astype(str).str.contains("Year", case=False, na=False).any():
            hdr = r; break
    if hdr is None:
        return pd.DataFrame(columns=COLS)
    df = pd.read_excel(_dp(fn), header=hdr)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    cols = list(df.columns)
    def pick(*keys):
        for c in cols:
            lc = str(c).lower()
            if all(k in lc for k in keys):
                return c
        return None
    c_year, c_la, c_phase = pick("year"), pick("la"), pick("phase")
    c_res = pick("mtd", "resident") or pick("resident", "headcoun")
    c_in, c_out = pick("in", "la"), pick("outside")
    c_lares = pick("la", "resident", "headcoun")
    c_ind = pick("independent")
    if not (c_year and c_la and c_phase):
        return pd.DataFrame(columns=COLS)
    out = pd.DataFrame({
        "year": df[c_year].map(lambda v: _year_start(v) if pd.notna(v) and str(v)[:4].isdigit() else np.nan),
        "year_label": df[c_year].map(lambda v: _academic_year_label(v) if pd.notna(v) and str(v)[:4].isdigit() else ""),
        "la": df[c_la].map(_norm_la),
        "phase": df[c_phase].astype(str).str.strip(),
        "mtd_resident": _num(df[c_res]) if c_res else np.nan,
        "mtd_in_la": _num(df[c_in]) if c_in else np.nan,
        "mtd_out_la": _num(df[c_out]) if c_out else np.nan,
        "la_resident": _num(df[c_lares]) if c_lares else np.nan,
        "indep_estimate": _num(df[c_ind]) if c_ind else np.nan,
    })
    return out.dropna(subset=["year"]).reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_internal_migration():
    """Net internal (domestic) migration by LA and age band, 2024.
    Prefers the small pre-aggregated CSV; falls back to aggregating ONS's full
    origin–destination workbook (large and slow, so the CSV is strongly preferred)."""
    COLS = ["area_code", "inflow", "outflow", "net", "age_band", "year"]
    small = "internal_migration_children_2024.csv"
    if _exists(small):
        df = pd.read_csv(_dp(small))
        for c in ["inflow", "outflow", "net"]:
            if c in df.columns:
                df[c] = _num(df[c])
        return df
    big = "detailedestimates2024on2023las.xlsx"
    if not _exists(big):
        return pd.DataFrame(columns=COLS)
    raw = pd.read_excel(_dp(big), sheet_name="IM2024 on 2023 LAs")
    bands = {"0-4": range(0, 5), "5-9": range(5, 10), "10-14": range(10, 15),
             "15-19": range(15, 20), "0-15": range(0, 16), "All ages": None}
    recs = []
    for band, rng in bands.items():
        cols = ([f"Age_{a}" for a in rng] if rng is not None
                else [c for c in raw.columns if str(c).startswith("Age_")])
        cols = [c for c in cols if c in raw.columns]
        if not cols:
            continue
        v = raw[cols].sum(axis=1)
        tmp = pd.DataFrame({"outla": raw["outla"], "inla": raw["inla"], "v": v})
        m = pd.concat([tmp.groupby("inla")["v"].sum().rename("inflow"),
                       tmp.groupby("outla")["v"].sum().rename("outflow")], axis=1).fillna(0.0)
        m["net"] = m["inflow"] - m["outflow"]; m["age_band"] = band; m["year"] = 2024
        m.index.name = "area_code"
        recs.append(m.reset_index())
    if not recs:
        return pd.DataFrame(columns=COLS)
    agg = pd.concat(recs, ignore_index=True)
    return agg[agg["area_code"].astype(str).str.match(r"E0[69]")].reset_index(drop=True)

# Nomis API endpoints from the data-sources note (used only if no local CSV exists)
NOMIS_BIRTHS_LA = ("https://www.nomisweb.co.uk/api/v01/dataset/NM_205_1.data.csv?"
                   "geography=1774190698,1774190704,1774190710,1774190711,1774190723,1774190724"
                   "&age_of_mother=101,103...109&measures=20100")
NOMIS_BIRTHS_LSOA = (
    "https://www.nomisweb.co.uk/api/v01/dataset/NM_206_1.data.csv?geography="
    "633344309,633344311,633344375,633344381,633344308,633344310,633344312,"
    "633344361,633344362,633344313,633344376,633344377,633344379,633344380,"
    "633344369...633344373,633344329,633344330,633344332...633344334,633344374,"
    "633344331,633344357...633344359,633344404,633344353,633344355,633344356,"
    "633344360,633373613,633344320,633344321,633344363,633344378,633344328,"
    "633344354,633371587,633371589...633371593,633344402,633344403,633344405,"
    "633370765,633370766,633344364,633344365,633344367,633344368,633344408,"
    "633344318,633344319,633344322...633344324,633344366,633344409,633344410,"
    "633373614,633344315...633344317,633344406,633373615,633344335,"
    "633344338...633344340,633371584,633371585,633344336,633344337,633344346,"
    "633344349,633344352,633344314,633344345,633344347,633344348,633344350,"
    "633344351,633344385...633344387,633344407,633344341...633344344,633373616,"
    "633344382...633344384,633344393,633344396,633344389,633344392,633344394,"
    "633344395,633373617,633373618,633344391,633344397,633344399...633344401,"
    "633344326,633344398,633371586,633371588,633373619,633344325,633344327,"
    "633344388,633344390,633373620"
    "&measures=20100")

NOMIS_POP_SYOA = ("https://www.nomisweb.co.uk/api/v01/dataset/NM_2002_1.data.csv?"
                  "geography=1774190698,1774190704,1774190710,1774190711,1774190723,1774190724"
                  "&gender=0&c_age=101...117&measures=20100")

@st.cache_data(show_spinner=False, ttl=86400)
def load_pop_syoa():
    """Resident population by SINGLE year of age, Westminster + neighbours (Nomis NM_2002_1).
    Used to build the school-phase cohorts described in the method note: Primary = ages 4–10
    (reception is age 4 turning 5; year 6 is age 10 turning 11) and Secondary = ages 11–16
    (all pupils turn 16 by the end of year 11)."""
    COLS = ["year", "la", "age", "population"]
    df = None
    for cand in ["population_single_year_of_age.csv", "NM_2002_1.data.csv", "pop_syoa.csv"]:
        if _exists(cand):
            df = pd.read_csv(_dp(cand), low_memory=False)
            break
    if df is None:
        df = _fetch_csv(NOMIS_POP_SYOA)
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS)
    cols = {str(c).upper(): c for c in df.columns}
    def find(*keys):
        for up, orig in cols.items():
            if all(k in up for k in keys):
                return orig
        return None
    c_geo = find("GEOGRAPHY", "NAME") or find("GEOGRAPHY")
    c_date = find("DATE", "NAME") or find("DATE")
    c_age = find("C_AGE", "NAME") or find("AGE", "NAME") or find("C_AGE")
    c_val = find("OBS_VALUE") or find("VALUE")
    if not (c_geo and c_date and c_age and c_val):
        return pd.DataFrame(columns=COLS)
    out = pd.DataFrame({
        "year": _num(df[c_date].astype(str).str.extract(r"(\d{4})")[0]),
        "la": df[c_geo].map(_norm_la),
        "age": _num(df[c_age].astype(str).str.extract(r"(\d+)")[0]),
        "population": _num(df[c_val]),
    })
    return out.dropna(subset=["year", "age", "population"]).reset_index(drop=True)

ETH_COL_MAP = {
    "ethnicity_white_british": "White British",
    "ethnicity_irish": "Irish",
    "ethnicity_traveller_of_irish_heritage": "Traveller of Irish Heritage",
    "ethnicity_any_other_white_background": "Any Other White Background",
    "ethnicity_gypsy_roma": "Gypsy/Roma",
    "ethnicity_white_and_black_caribbean": "White and Black Caribbean",
    "ethnicity_white_and_black_african": "White and Black African",
    "ethnicity_white_and_asian": "White and Asian",
    "ethnicity_any_other_mixed_background": "Any Other Mixed Background",
    "ethnicity_indian": "Indian",
    "ethnicity_pakistani": "Pakistani",
    "ethnicity_bangladeshi": "Bangladeshi",
    "ethnicity_any_other_asian_background": "Any Other Asian Background",
    "ethnicity_black_caribbean": "Black Caribbean",
    "ethnicity_black_african": "Black African",
    "ethnicity_any_other_black_background": "Any Other Black Background",
    "ethnicity_chinese": "Chinese",
    "ethnicity_any_other_ethnic_group": "Any Other Ethnic Group",
    "ethnicity_unclassified": "Unclassified",
}
ETH_HIGH_MAP = {
    "White British": "White", "Irish": "White", "Traveller of Irish Heritage": "White",
    "Any Other White Background": "White", "Gypsy/Roma": "White",
    "White and Black Caribbean": "Mixed", "White and Black African": "Mixed",
    "White and Asian": "Mixed", "Any Other Mixed Background": "Mixed",
    "Indian": "Asian", "Pakistani": "Asian", "Bangladeshi": "Asian",
    "Any Other Asian Background": "Asian",
    "Black Caribbean": "Black", "Black African": "Black", "Any Other Black Background": "Black",
    "Chinese": "Other", "Any Other Ethnic Group": "Other", "Unclassified": "Unclassified",
}

@st.cache_data(show_spinner=False)
def load_sen_ethnicity():
    """DfE SEN/FSM/language by ethnicity - London region + all 33 London boroughs.
    Source file: sen_fsm_eth_lang_new_.csv (has la_name/new_la_code for LA rows,
    region_name='London' for the regional row). Returns tidy LONG:
    year, year_label, geo, phase, provision, ethnicity, eth_high, count."""
    COLS = ["year", "year_label", "geo", "phase", "provision", "ethnicity", "eth_high", "count"]
    fn = "sen_fsm_eth_lang_new_.csv"
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    df = pd.read_csv(_dp(fn), low_memory=False)
    lvl = df["geographic_level"].astype(str).str.lower()
    geo = np.where(lvl == "regional", df["region_name"].astype(str), df["la_name"].astype(str))
    base = pd.DataFrame({
        "year": df["time_period"].map(_year_start),
        "year_label": df["time_period"].map(_academic_year_label),
        "geo": pd.Series(geo).map(_norm_la),
        "phase": df["phase_type_grouping"].astype(str).str.strip(),
        "provision": df["sen_status"].astype(str).str.strip(),
        "sen_need": df["sen_primary_need"].astype(str).str.strip(),
    })
    eth_cols = [c for c in df.columns if c in ETH_COL_MAP]
    frames = []
    for c in eth_cols:
        f = base.copy()
        f["ethnicity"] = ETH_COL_MAP[c]
        f["eth_high"] = ETH_HIGH_MAP.get(ETH_COL_MAP[c], "Other")
        f["count"] = _num(df[c])
        frames.append(f)
    out = pd.concat(frames, ignore_index=True)
    # collapse across SEN primary need so counts aren't multiplied across need types
    out = out[out["sen_need"] == "Total"].drop(columns=["sen_need"])
    out = out[out["geo"].astype(str).str.lower() != "nan"]
    return out.dropna(subset=["count"]).reset_index(drop=True)[COLS]

@st.cache_data(show_spinner=False)
def load_fsm_language():
    """FSM eligibility and language (English/other) from the same source file,
    for the FSM/language callouts alongside the SEN-by-ethnicity chart.
    Returns LONG: year, year_label, geo, phase, provision, metric, count."""
    COLS = ["year", "year_label", "geo", "phase", "provision", "metric", "count"]
    fn = "sen_fsm_eth_lang_new_.csv"
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    df = pd.read_csv(_dp(fn), low_memory=False)
    lvl = df["geographic_level"].astype(str).str.lower()
    geo = np.where(lvl == "regional", df["region_name"].astype(str), df["la_name"].astype(str))
    base = pd.DataFrame({
        "year": df["time_period"].map(_year_start),
        "year_label": df["time_period"].map(_academic_year_label),
        "geo": pd.Series(geo).map(_norm_la),
        "phase": df["phase_type_grouping"].astype(str).str.strip(),
        "provision": df["sen_status"].astype(str).str.strip(),
        "sen_need": df["sen_primary_need"].astype(str).str.strip(),
    })
    METRIC_MAP = {"fsm_eligible": "FSM eligible", "fsm_not_eligible": "Not FSM eligible",
                  "language_english": "English as first language",
                  "language_other": "Other first language",
                  "language_unclassified": "Language unclassified"}
    frames = []
    for c, label in METRIC_MAP.items():
        if c in df.columns:
            f = base.copy()
            f["metric"] = label
            f["count"] = _num(df[c])
            frames.append(f)
    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=COLS)
    if out.empty:
        return out
    out = out[out["sen_need"] == "Total"].drop(columns=["sen_need"])
    out = out[out["geo"].astype(str).str.lower() != "nan"]
    return out.dropna(subset=["count"]).reset_index(drop=True)[COLS]

NOMIS_FERTILITY = (
    "https://www.nomisweb.co.uk/api/v01/dataset/NM_207_1.data.csv?geography="
    "1774190693...1774190698,1774190692,1774190699...1774190724,"
    "2092957699,2013265927"
    "&measure=2...4,12,5...11&measures=20100")

# ── Fertility rates (CBR · GFR · TFR) ─────────────────────────────────────────
# The data-sources note asks for crude birth rate, general fertility rate and total
# fertility rate for the nearest neighbours, London and England, but gives no link.
# Two routes are supported: a ready-made rates table, or derivation from births by
# age of mother (NM_205_1) plus female population by age band.
def _age_band_bounds(label):
    """'Aged 20-24' -> (20,24); 'Under 20' -> (15,19); '45 and over' -> (45,49)."""
    s = str(label).lower()
    nums = [int(n) for n in re.findall(r"\d+", s)]
    if "under" in s and nums:
        return (15, nums[0] - 1)
    if ("over" in s or "plus" in s or "+" in s) and nums:
        return (nums[0], 49)
    if len(nums) >= 2:
        return (nums[0], nums[1])
    if len(nums) == 1:
        return (nums[0], nums[0])
    return None

@st.cache_data(show_spinner=False)
def load_fertility_rates():
    """Fertility and birth rates (TFR, GFR, crude birth rate, age-specific rates) for
    London boroughs, London and England — Nomis dataset NM_207_1, or a local CSV.

    Nomis CSVs carry BOTH a numeric code column and a text label column for every
    dimension (`MEASURE` and `MEASURE_NAME`, `GEOGRAPHY` and `GEOGRAPHY_NAME`), plus an
    unrelated `MEASURES_NAME` field. Matching on a bare substring picks the numeric code,
    so `_nomis_pick` explicitly prefers the `*_NAME` label column."""
    COLS = ["year", "area", "measure", "value"]
    df = None
    for cand in ["fertility_rates.csv", "NM_207_1.data.csv",
                 "births_fertility_rates.csv", "tfr_gfr_cbr.csv"]:
        if _exists(cand):
            df = pd.read_csv(_dp(cand), low_memory=False)
            break
    if df is None:
        df = _fetch_csv(NOMIS_FERTILITY, timeout=40)
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS)

    cols = list(df.columns)
    def _nomis_pick(stem, *fallbacks):
        """Prefer '<stem>_name', then an exact '<stem>', then the fallbacks."""
        low = {str(c).lower(): c for c in cols}
        if f"{stem}_name" in low:
            return low[f"{stem}_name"]
        if stem in low:
            return low[stem]
        for fb in fallbacks:
            for c in cols:
                lc = str(c).lower()
                if fb in lc and not lc.startswith("measures"):
                    return c
        return None

    c_year = _nomis_pick("date", "year", "period")
    c_area = _nomis_pick("geography", "area", "la_name")
    c_meas = _nomis_pick("measure", "indicator", "rate_type", "metric")
    c_val = _nomis_pick("obs_value", "value", "rate")
    if not (c_year and c_area and c_val):
        return pd.DataFrame(columns=COLS)

    out = pd.DataFrame({
        "year": _num(df[c_year].astype(str).str.extract(r"(\d{4})")[0]),
        "area": df[c_area].map(_norm_la),
        "measure": (df[c_meas].astype(str).str.strip() if c_meas is not None else "Rate"),
        "value": _num(df[c_val]),
    })
    # a numeric 'measure' means we picked the code column, not the label — unusable
    if out["measure"].str.fullmatch(r"\d+(\.0)?").fillna(False).all():
        out["measure"] = "Rate"
    return out.dropna(subset=["year", "value"]).reset_index(drop=True)

def order_fertility_measures(measures):
    """Headline rates first, then age-specific rates, then everything else."""
    def rank(m):
        s = str(m).lower()
        if "total fertility" in s or s.strip() in ("tfr",):
            return (0, s)
        if "general fertility" in s or s.strip() in ("gfr",):
            return (1, s)
        if "crude birth" in s or s.strip() in ("cbr",):
            return (2, s)
        if "mean age" in s:
            return (4, s)
        if any(ch.isdigit() for ch in s):
            return (3, s)
        return (5, s)
    return sorted(measures, key=rank)

@st.cache_data(show_spinner=False)
def load_female_pop():
    """Female population by age band, used to derive fertility rates.
    Expected: year · area · age band · population (any Nomis/ONS extract shape)."""
    COLS = ["year", "area", "age_band", "population"]
    fn = None
    for cand in ["female_population_by_age.csv", "female_pop.csv",
                 "population_females_by_age.csv"]:
        if _exists(cand):
            fn = cand; break
    if fn is None:
        return pd.DataFrame(columns=COLS)
    df = pd.read_csv(_dp(fn), low_memory=False)
    cols = list(df.columns)
    def pick(*keys):
        for c in cols:
            if all(k in str(c).lower() for k in keys):
                return c
        return None
    c_year = pick("date", "name") or pick("year") or pick("date")
    c_area = pick("geography", "name") or pick("area") or pick("geography")
    c_age = pick("c_age", "name") or pick("age", "name") or pick("age")
    c_val = pick("obs_value") or pick("value") or pick("population")
    if not (c_year and c_area and c_age and c_val):
        return pd.DataFrame(columns=COLS)
    out = pd.DataFrame({
        "year": _num(df[c_year].astype(str).str.extract(r"(\d{4})")[0]),
        "area": df[c_area].map(_norm_la),
        "age_band": df[c_age].astype(str).str.strip(),
        "population": _num(df[c_val]),
    })
    return out.dropna(subset=["year", "population"]).reset_index(drop=True)

def derive_fertility_rates(df_births, df_fpop):
    """Derive GFR and TFR from births by age of mother and female population by band.

        ASFR(band) = births to mothers in band / female population in band
        GFR        = total births / females aged 15–44 × 1000
        TFR        = Σ (ASFR × band width)

    Bands are matched on their numeric bounds, so Nomis and ONS labels can differ."""
    COLS = ["year", "area", "measure", "value"]
    if df_births is None or df_births.empty or df_fpop is None or df_fpop.empty:
        return pd.DataFrame(columns=COLS)
    b = df_births.copy()
    b["bounds"] = b["age_of_mother"].map(_age_band_bounds)
    b = b[b["bounds"].notna()]
    f = df_fpop.copy()
    f["bounds"] = f["age_band"].map(_age_band_bounds)
    f = f[f["bounds"].notna()]
    if b.empty or f.empty:
        return pd.DataFrame(columns=COLS)
    b = b.groupby(["year", "la", "bounds"], as_index=False)["births"].sum().rename(columns={"la": "area"})
    f = f.groupby(["year", "area", "bounds"], as_index=False)["population"].sum()
    m = b.merge(f, on=["year", "area", "bounds"], how="inner")
    if m.empty:
        return pd.DataFrame(columns=COLS)
    m = m[m["population"] > 0].copy()
    m["width"] = m["bounds"].map(lambda t: t[1] - t[0] + 1)
    m["asfr"] = m["births"] / m["population"]
    out = []
    for (yr, area), g in m.groupby(["year", "area"]):
        tfr = float((g["asfr"] * g["width"]).sum())
        rep = g[g["bounds"].map(lambda t: t[0] >= 15 and t[1] <= 44)]
        gfr = (float(rep["births"].sum()) / float(rep["population"].sum()) * 1000
               if rep["population"].sum() > 0 else np.nan)
        out.append((int(yr), area, "Total fertility rate (TFR)", tfr))
        if pd.notna(gfr):
            out.append((int(yr), area, "General fertility rate (GFR)", gfr))
    return pd.DataFrame(out, columns=COLS)

SCHOOL_COHORTS = {"Primary cohort (ages 4–10)": (4, 10),

                  "Secondary cohort (ages 11–16)": (11, 16),
                  "Year 6→7 transition (ages 10–12)": (10, 12),
                  "Year 11→12 transition (ages 15–17)": (15, 17)}

# Age ranges behind each school phase, per the council's method note:
# reception is age 4 turning 5 and year 6 is age 10 turning 11; every secondary
# pupil turns 16 by the end of year 11.
PHASE_AGES = {"Primary": (4, 10), "Secondary": (11, 16)}

@st.cache_data(show_spinner=False)
def load_mye_syoa_wcc():
    """Westminster resident population by SINGLE year of age, from the small-area MYE
    workbook (which holds F0..F90 / M0..M90 columns). Covers mid-2022 to mid-2024 and
    Westminster only — used as a fallback for the independent-school calculation when
    the Nomis single-year-of-age extract (all six boroughs) has not been supplied."""
    fn = "Small_Area_Output_Area_Mid_Year_Estimated.xlsx"
    COLS = ["year", "la", "age", "population"]
    if not _exists(fn):
        return pd.DataFrame(columns=COLS)
    sheets = {"Mid-2022 LSOA 2021": 2022, "Mid-2023 LSOA 2021": 2023, "Mid-2024 LSOA 2021": 2024}
    recs = []
    for sheet, yr in sheets.items():
        try:
            d = pd.read_excel(_dp(fn), sheet_name=sheet, header=3)
        except Exception:
            continue
        if "LAD 2023 Name" not in d.columns:
            continue
        d = d[d["LAD 2023 Name"] == "Westminster"]
        for a in range(0, 21):
            f, m = f"F{a}", f"M{a}"
            if f in d.columns and m in d.columns:
                recs.append((yr, "Westminster", a, float(d[f].sum() + d[m].sum())))
    return pd.DataFrame(recs, columns=COLS)

def cohort_population(la, lo, hi, df_syoa, df_wcc):
    """Resident population aged lo–hi by MYE year for one local authority.
    Prefers the Nomis single-year-of-age extract (all boroughs); falls back to the
    Westminster-only small-area workbook. Returns year → population."""
    if df_syoa is not None and not df_syoa.empty:
        d = df_syoa[(df_syoa["la"] == la) & (df_syoa["age"] >= lo) & (df_syoa["age"] <= hi)]
        if not d.empty:
            g = d.groupby("year", as_index=False)["population"].sum()
            return dict(zip(g["year"].astype(int), g["population"])), "Nomis single year of age"
    if df_wcc is not None and not df_wcc.empty and la == "Westminster":
        d = df_wcc[(df_wcc["age"] >= lo) & (df_wcc["age"] <= hi)]
        if not d.empty:
            g = d.groupby("year", as_index=False)["population"].sum()
            return dict(zip(g["year"].astype(int), g["population"])), "ONS small-area MYE (Westminster only)"
    return {}, None

def compute_independent_estimate(df_xb, df_syoa, df_wcc):
    """Independent-school estimate, calculated rather than read from a spreadsheet:

        independent = LA resident headcount (ONS MYE, phase age range)
                      − maintained-school resident headcount (DfE cross-border)
        % independent = independent / LA resident headcount

    An academic year is matched to the mid-year estimate at its END (2023/24 → mid-2024),
    which is the mapping that reproduces the council's own reference table."""
    COLS = ["year", "year_label", "la", "phase", "mye_year", "la_resident",
            "mtd_resident", "indep_estimate", "pct_independent", "source"]
    if df_xb is None or df_xb.empty:
        return pd.DataFrame(columns=COLS)
    out = []
    for la in sorted(df_xb["la"].unique()):
        for phase, (lo, hi) in PHASE_AGES.items():
            pops, src = cohort_population(la, lo, hi, df_syoa, df_wcc)
            if not pops:
                continue
            sub = df_xb[(df_xb["la"] == la) & (df_xb["phase"].str.lower() == phase.lower())]
            for _, r in sub.iterrows():
                mye_year = int(r["year"]) + 1          # academic 2023/24 → mid-2024
                pop = pops.get(mye_year)
                if pop is None or not pop or pd.isna(r["resident"]):
                    continue
                ind = pop - float(r["resident"])
                out.append((int(r["year"]), r["year_label"], la, phase, mye_year, pop,
                            float(r["resident"]), ind, ind / pop * 100, src))
    return pd.DataFrame(out, columns=COLS)

@st.cache_data(show_spinner=False, ttl=86400)
def load_births_la():
    """Live births by LA and age of mother. Reads a local CSV from data/ if present,
    otherwise calls the Nomis API. Returns tidy: year, la, age_of_mother, births."""
    COLS = ["year", "la", "age_of_mother", "births"]
    df = None
    for cand in ["births_by_age_of_mother.csv", "NM_205_1.data.csv", "births_la.csv"]:
        if _exists(cand):
            df = pd.read_csv(_dp(cand), low_memory=False)
            break
    if df is None:
        df = _fetch_csv(NOMIS_BIRTHS_LA)
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS)
    cols = {str(c).upper(): c for c in df.columns}
    def find(*keys):
        for up, orig in cols.items():
            if all(k in up for k in keys):
                return orig
        return None
    c_geo = find("GEOGRAPHY", "NAME") or find("GEOGRAPHY")
    c_date = find("DATE", "NAME") or find("DATE")
    c_age = find("AGE_OF_MOTHER", "NAME") or find("AGE_OF_MOTHER")
    c_val = find("OBS_VALUE") or find("VALUE")
    if not (c_geo and c_date and c_val):
        return pd.DataFrame(columns=COLS)
    out = pd.DataFrame({
        "year": _num(df[c_date].astype(str).str.extract(r"(\d{4})")[0]),
        "la": df[c_geo].map(_norm_la),
        "age_of_mother": df[c_age].astype(str).str.strip() if c_age else "All ages",
        "births": _num(df[c_val]),
    })
    return out.dropna(subset=["year", "births"]).reset_index(drop=True)

@st.cache_data(show_spinner=False, ttl=86400)
def load_births_lsoa():
    """Live births by Westminster LSOA over time (Nomis NM_206_1 or a local CSV).
    Returns tidy: year, LSOA_CODE, LSOA_NAME, births."""
    COLS = ["year", "LSOA_CODE", "LSOA_NAME", "births"]
    df = None
    for cand in ["births_by_lsoa.csv", "NM_206_1.data.csv", "births_lsoa.csv"]:
        if _exists(cand):
            df = pd.read_csv(_dp(cand), low_memory=False)
            break
    if df is None:
        df = _fetch_csv(NOMIS_BIRTHS_LSOA, timeout=40)   # ~120 Westminster LSOAs
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS)
    cols = {str(c).upper(): c for c in df.columns}
    def find(*keys):
        for up, orig in cols.items():
            if all(k in up for k in keys):
                return orig
        return None
    c_name = find("GEOGRAPHY", "NAME") or find("GEOGRAPHY")
    c_code = find("GEOGRAPHY", "CODE")
    c_date = find("DATE", "NAME") or find("DATE")
    c_val = find("OBS_VALUE") or find("VALUE")
    if not (c_name and c_date and c_val):
        return pd.DataFrame(columns=COLS)
    codes = (df[c_code].astype(str).str.extract(r"(E\d{8,})")[0] if c_code
             else _lsoa_code(df[c_name]))
    out = pd.DataFrame({
        "year": _num(df[c_date].astype(str).str.extract(r"(\d{4})")[0]),
        "LSOA_CODE": codes,
        "LSOA_NAME": _lsoa_name(df[c_name]),
        "births": _num(df[c_val]),
    })
    return out.dropna(subset=["year", "births", "LSOA_CODE"]).reset_index(drop=True)

# ── derived metrics shared by the schools charts ──────────────────────────────
def sen_totals(df_sen, phase):
    """Total pupil headcount by geography × year for a phase (or all phases)."""
    if df_sen.empty:
        return pd.DataFrame(columns=["year", "year_label", "geo", "count"])
    d = df_sen[df_sen["provision"] == "Total"]
    if phase != "All phases":
        d = d[d["phase"] == phase]
    return d.groupby(["year", "year_label", "geo"], as_index=False)["count"].sum()

def index_to_baseline(d, value_col="count", base_year=None, mode="pct_change"):
    """% change (or index=100) against each geography's own baseline year."""
    if d.empty:
        return d
    out = d.copy()
    base_year = base_year if base_year is not None else out["year"].min()
    base = out[out["year"] == base_year].set_index("geo")[value_col]
    out["_base"] = out["geo"].map(base)
    out = out[out["_base"].notna() & (out["_base"] != 0)]
    out["value"] = ((out[value_col] / out["_base"]) - 1) * 100 if mode == "pct_change" \
        else (out[value_col] / out["_base"]) * 100
    return out.drop(columns=["_base"])

# ══════════════════════════════════════════════════════════════════════════════
# LOAD EVERYTHING  — every loader is isolated so one missing/corrupt file can
# never crash the whole dashboard; it just disables the sections that need it.
# ══════════════════════════════════════════════════════════════════════════════
def _diagnose_file(name):
    """Explain *why* a data file could not be read (esp. a corrupt/LFS .xlsx)."""
    p = _dp(name)
    if not os.path.exists(p):
        return f"`{name}` was not found on the server."
    try:
        with open(p, "rb") as fh:
            head = fh.read(64)
    except Exception:
        return f"`{name}` could not be opened."
    if head.startswith(b"version https://git-lfs"):
        return (f"`{name}` is only a **Git LFS pointer** on the server, not the real "
                "spreadsheet — the binary wasn't pulled at deploy time. Commit the actual "
                "file (or fetch LFS objects), then redeploy.")
    if name.lower().endswith((".xlsx", ".xlsm")) and head[:2] != b"PK":
        return (f"`{name}` on the server is **not a valid .xlsx** (its bytes don't start with "
                "the ZIP signature). This almost always means the file was corrupted when it "
                "was committed to Git as if it were text. Add a `.gitattributes` line "
                "`*.xlsx binary`, re-commit the file (or re-upload it through the GitHub web "
                "interface), then redeploy.")
    return None

def _safe_load(loader, default, probe=None):
    """Run a loader; on any error keep the app alive and surface a clear reason."""
    try:
        return loader()
    except Exception as exc:
        label = loader.__name__.replace("load_", "").replace("_", " ")
        reason = _diagnose_file(probe) if probe else None
        st.warning(f"⚠️ Could not load the **{label}** dataset, so its charts are hidden. "
                   + (reason or f"({type(exc).__name__}: {exc})"))
        return default

with st.spinner("Loading datasets…"):
    df_mye_la    = _safe_load(load_mye_la,    pd.DataFrame(), "MYEs_LA_1991_2024_gender.xlsx")
    df_mye_lsoa  = _safe_load(load_mye_lsoa,  pd.DataFrame(), "Small_Area_Output_Area_Mid_Year_Estimated.xlsx")
    df_imd       = _safe_load(load_imd,       pd.DataFrame(), "File_1_IoD2025_Index_of_Multiple_Deprivation.xlsx")
    df_idaci     = _safe_load(load_idaci,     pd.DataFrame(), "File_3_IoD2025_Supplementary_Indices_IDACI_and_IDAOPI.xlsx")
    ward_gj      = _safe_load(load_ward_geojson, None,        "Wards_WCC.json")
    ward_lookup  = _safe_load(load_ward_lookup,  pd.DataFrame(columns=["LSOA_CODE", "Ward", "coverage"]), "Ward LSOA Lookup.xlsx")
    lsoa_to_ward = ward_for_lsoa(ward_lookup)
    df_li_la     = _safe_load(load_low_income_la,   pd.DataFrame(), "2_AHC_Relative_LA.csv")
    df_li_ward   = _safe_load(load_low_income_ward, pd.DataFrame(), "4_AHC_Relative_Ward.csv")
    df_ks4_eth   = _safe_load(load_ks4_ethnic, pd.DataFrame(), "data-key-stage-4-performance.csv")
    df_ks4_ts    = _safe_load(load_ks4_time,   pd.DataFrame(), "data-key-stage-4-performance__1_.csv")
    df_rm006     = _safe_load(load_rm006, pd.DataFrame(), "RM006_age_of_youngest_dependent_child_by_household_type.xlsx")
    df_rm012     = _safe_load(load_rm012, pd.DataFrame(), "RM012_dependent_children_by_ethnic_group_of_HRP.xlsx")
    df_rm033     = _safe_load(load_rm033, pd.DataFrame(), "RM033_ethic_group_dependent_child_by_sex.xlsx")
    df_egdi      = _safe_load(load_egdi,      pd.DataFrame(), "EGDI-Local-Authority-profiles.xlsx")
    df_egdi_lsoa = _safe_load(load_egdi_lsoa, pd.DataFrame(), "EGDI.xlsx")
    lsoa_gj      = _safe_load(load_lsoa_geojson,    None, "ONS_LSOA_2021 (1).json")
    borough_gj   = _safe_load(load_borough_geojson, None, "Borough_London_LL84.json")
    # ── new datasets: children & schools, population change, childcare costs
    df_pupils    = _safe_load(load_pupils, pd.DataFrame(), "data-school-pupils-and-their-characteristics.csv")
    df_sen       = _safe_load(load_sen,    pd.DataFrame(), "SEN_data-special-educational-needs-in-england.csv")
    _xb          = _safe_load(load_cross_border, (pd.DataFrame(), pd.DataFrame()),
                              "cross_border_data_data-school-pupils-and-their-characteristics.csv")
    df_xborder, df_xflows = _xb if isinstance(_xb, tuple) else (pd.DataFrame(), pd.DataFrame())
    df_ccosts    = _safe_load(load_childcare_costs, pd.DataFrame(), "costs_data-childcare-and-early-years-provider-survey.csv")
    df_indep     = _safe_load(load_indep_calc, pd.DataFrame(), "Independent_schools_calculations.xlsx")
    df_migr      = _safe_load(load_internal_migration, pd.DataFrame(), "internal_migration_children_2024.csv")
    df_births    = _safe_load(load_births_la, pd.DataFrame())
    df_births_lsoa = _safe_load(load_births_lsoa, pd.DataFrame())
    df_syoa      = _safe_load(load_pop_syoa, pd.DataFrame())
    df_wcc_syoa  = _safe_load(load_mye_syoa_wcc, pd.DataFrame(), "Small_Area_Output_Area_Mid_Year_Estimated.xlsx")
    df_indep_calc = compute_independent_estimate(df_xborder, df_syoa, df_wcc_syoa)
    df_sen_eth   = _safe_load(load_sen_ethnicity, pd.DataFrame(), "sen_fsm_eth_lang_new_.csv")
    df_fsm_lang  = _safe_load(load_fsm_language,  pd.DataFrame(), "sen_fsm_eth_lang_new_.csv")
    df_fert      = _safe_load(load_fertility_rates, pd.DataFrame())
    df_fpop      = _safe_load(load_female_pop, pd.DataFrame())

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
st.title("🏙️ WCC Children's Demographics")
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
tab0, tab7, tab2, tab1, tab5, tab6, tab3, tab4 = st.tabs([
    "🏠 Overview",
    "📉 Births, Migration & Decline",
    "🗺️ Population & Demographics",
    "📍 Child Poverty",
    "🎒 Children & Schools",
    "🧸 Childcare Costs",
    "📚 KS4 Attainment",
    "⚖️ Deprivation (IMD · IDACI · EGDI)",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — OVERVIEW / LANDING
# ══════════════════════════════════════════════════════════════════════════════
with tab0:
    st.subheader("Datasets - reliability and relevance")
    st.markdown(
        "Westminster's child population is measured by several official sources, each built "
        "differently and each best for a particular question. This page explains what to use when, "
        "so the numbers later in the dashboard are read in the right context.")

    st.markdown("##### The datasets at a glance")
    st.markdown(f"""
<div class="ds-card"><b>① Mid-Year Estimates (MYEs) — start here.</b><br>
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
child poverty geography</b> — it measures the share of children in income-deprived families and, the
IMD's 2025 update includes after-housing-costs, exposing hidden child poverty in high-cost areas like Westminster.
<b>Prefer the IMD alongside IDACI.</b></div>

<div class="ds-card"><b>④ Ethnic Group Deprivation Index (EGDI) — the ethnicity dimension of deprivation.</b><br>
IMD and IDACI tell you <i>where</i> deprivation sits; they cannot tell you whether it falls
<b>unevenly across ethnic groups</b> within the same neighbourhood. The EGDI adds exactly that lens.
Reading it alongside IDACI and IMD turns "this area is deprived" into "and deprivation here is borne
disproportionately by particular ethnic groups" — essential for targeting support equitably.</div>

<div class="ds-card"><b>⑤ DfE schools data (pupils, SEN, cross-border) — the school-age reality.</b><br>
Where the MYEs count <i>resident</i> children, the DfE data counts <b>pupils on rolls</b> — and the two
diverge sharply in Westminster because so many children are educated privately or outside the borough.
The SEN publication is doubly useful: alongside SEN counts it reports the <b>total headcount</b> for
every phase nationally, regionally and for all 33 London boroughs, which is what lets us benchmark
Westminster against the <b>London and England averages</b>.</div>

<div class="ds-card"><b>⑥ Births, internal migration and childcare costs — the drivers.</b><br>
Births set the size of each future school cohort; <b>internal migration</b> shows families leaving
(and students arriving); <b>childcare costs</b> are one of the pressures behind that movement. Read
together they explain <i>why</i> the child population is falling, rather than just showing that it is.</div>
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
                    tr.line.width = 2; tr.marker.size = 6   # distinct colour kept (from palette)
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
        _bpal = borough_palette(list(NEIGHBOURS))
        for b in NEIGHBOURS:
            sb = dla[dla["area"].str.contains(b.replace("& ", "").split()[0], case=False, na=False)]
            sb = dla[dla["area"] == b] if (dla["area"] == b).any() else sb
            ser = sb.groupby("year", as_index=False)["population"].sum().sort_values("year")
            if ser.empty:
                continue
            focal = (b == "Westminster")
            figt.add_trace(go.Scatter(
                x=ser["year"], y=ser["population"], mode="lines", name=b,
                line=dict(color=_bpal[b], width=3.8 if focal else 1.8),
                opacity=1.0 if focal else 0.95,
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
    st.subheader("Key Stage 4 attainment — Westminster vs CIPFA neighbours")
    st.markdown(
        "**Dataset:** DfE *Key Stage 4 performance*, accessed via Explore Education Statistics. "
        "**Attainment 8** is a pupil's average grade across eight core GCSE subjects (max 90). Here it "
        "is shown for Westminster against its CIPFA neighbours, by ethnic group, over time, and across "
        "boroughs. In the borough-comparison charts each neighbour keeps its own colour (consistent with "
        "the other tabs); Westminster is always strongest.")
    legend_hint()

    _ALL_ETH = ("total", "all", "all pupils", "all ethnic groups")
    if df_ks4_eth.empty:
        st.info("KS4 ethnicity file `data-key-stage-4-performance.csv` not found in the data folder.")
    else:
        # Shared ethnicity selector — drives BOTH the borough bar and the borough map
        specific = sorted(g for g in df_ks4_eth["ethnic_group"].dropna().unique()
                          if str(g).lower() not in _ALL_ETH)
        eth_opts = ["Average (all ethnicities)"] + specific
        sel_eth = st.selectbox("Ethnic group (applies to the bar and the map below)",
                               eth_opts, key="ks4_eth_sel")

        def _ks4_base(df):
            if sel_eth == "Average (all ethnicities)":
                tot = df[df["ethnic_group"].astype(str).str.lower().isin(_ALL_ETH)]
                src = tot if not tot.empty else df
            else:
                src = df[df["ethnic_group"] == sel_eth]
            return src.dropna(subset=["att8_2425"]).groupby("la", as_index=False)["att8_2425"].mean()

        base = _ks4_base(df_ks4_eth)
        cipfa = list(NEIGHBOURS)                      # CIPFA borough names (with & )
        bpal = borough_palette(cipfa)

        # ── Borough comparison bar — CIPFA neighbours, each a distinct muted colour
        bar = base[base["la"].isin(cipfa)].copy().sort_values("att8_2425")
        if not bar.empty:
            chart_title(f"Attainment 8 across CIPFA neighbours — {sel_eth.lower()} (2024/25)",
                        "Average Attainment 8 score · each borough its own colour · Westminster strongest")
            figc = go.Figure(go.Bar(
                x=bar["att8_2425"], y=bar["la"], orientation="h",
                marker_color=[bpal.get(l, CONTEXT_BAR) for l in bar["la"]],
                text=[f"{v:.1f}" for v in bar["att8_2425"]], textposition="outside",
                hovertemplate="<b>%{y}</b><br>Attainment 8: %{x:.1f}<extra></extra>"))
            figc.update_xaxes(title="Average Attainment 8 score")
            figc.update_yaxes(title="")
            figc.update_layout(height=360)
            show_chart(figc, "ks4_cipfa", "DfE KS4 performance, 2024/25")

        # ── Borough map — same ethnicity selection, kept in the blue map scheme
        if borough_gj is not None and not base.empty:
            id_by_name = {}
            for ft in borough_gj["features"]:
                p = ft["properties"]
                nm = (p.get("name") or p.get("BoroughNa") or p.get("NAME") or
                      next((v for k, v in p.items() if isinstance(v, str) and "E09" not in v), ""))
                id_by_name[str(nm).replace("and Fulham", "& Fulham").replace("and Chelsea", "& Chelsea")] = ft["id"]
            mp = base.copy()
            mp["gid"] = mp["la"].map(id_by_name)
            geo = mp.dropna(subset=["gid"])
            if not geo.empty:
                chart_title(f"Attainment 8 across inner-London boroughs — {sel_eth.lower()} (2024/25)",
                            "DfE KS4 · darker = higher Attainment 8 · hover for the score")
                figm = choropleth(borough_gj, geo["gid"], geo["att8_2425"], geo["la"],
                                  "Attainment 8", [[0, "#EDEFF6"], [1, FOCAL]], fmt=":.1f",
                                  zoom=9.2, center={"lat": 51.51, "lon": -0.12}, height=520)
                show_chart(figm, "ks4_map", "DfE KS4 performance, 2024/25")

        # ── Within Westminster, by ethnic group (single-borough, blue)
        wcc_e = df_ks4_eth[df_ks4_eth["la"].astype(str).str.contains("Westminster", na=False)].copy()
        wcc_e = wcc_e[~wcc_e["ethnic_group"].astype(str).str.lower().isin(_ALL_ETH)]
        wcc_e = wcc_e.dropna(subset=["att8_2425"]).sort_values("att8_2425")
        if not wcc_e.empty:
            chart_title("Within Westminster, Attainment 8 varies widely by ethnic group (2024/25)",
                        "Average Attainment 8 score, Westminster pupils · DfE KS4")
            fige = go.Figure(go.Bar(
                x=wcc_e["att8_2425"], y=wcc_e["ethnic_group"], orientation="h",
                marker_color=FOCAL, text=[f"{v:.1f}" for v in wcc_e["att8_2425"]],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Attainment 8: %{x:.1f}<extra></extra>"))
            fige.update_xaxes(title="Average Attainment 8 score")
            fige.update_yaxes(title="")
            show_chart(fige, "ks4_ethnic", "DfE KS4 performance, 2024/25")

    # ── Attainment 8 over time — CIPFA neighbours each in a distinct muted colour
    if df_ks4_ts.empty:
        st.info("KS4 time-series file `data-key-stage-4-performance__1_.csv` not found.")
    else:
        chart_title("Attainment 8 over time — Westminster vs CIPFA neighbours",
                    "Average Attainment 8 score · each borough its own colour · Westminster strongest")
        cipfa = list(NEIGHBOURS)
        tpal = borough_palette(cipfa)
        figt = go.Figure()
        for la in cipfa:
            s = df_ks4_ts[df_ks4_ts["la"] == la].dropna(subset=["att8"]).sort_values("year")
            if s.empty:
                continue
            focal = (la == "Westminster")
            figt.add_trace(go.Scatter(
                x=s["year"], y=s["att8"], mode="lines+markers", name=la,
                line=dict(color=tpal[la], width=3.8 if focal else 1.8),
                marker=dict(size=8 if focal else 5),
                hovertemplate="<b>"+la+"</b><br>%{x}: %{y:.1f}<extra></extra>"))
        figt.update_xaxes(title="Academic year")
        figt.update_yaxes(title="Average Attainment 8 score")
        show_chart(figt, "ks4_trend", "DfE KS4 performance, 2018/19–2024/25")
        legend_hint("Hide boroughs from the legend to read Westminster's trend on its own.")
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
                        "EGDI range = gap between the most- and least-deprived ethnic group in the LSOA · top three highlighted · ward shown on hover")
            lbl = np.where(er["Ward"] != "—", er["LSOA_NAME"] + " · " + er["Ward"], er["LSOA_NAME"])
            top3_cut = er["Range"].nlargest(3).min()          # highlight the three widest gaps only
            bar_cols = np.where(er["Range"] >= top3_cut, FOCAL, CONTEXT_BAR)
            figr = go.Figure(go.Bar(
                x=er["Range"], y=lbl, orientation="h", marker_color=bar_cols,
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

    try:
        prof = _imd_decile_profile()
    except Exception:
        prof = pd.DataFrame()
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
        rpal = borough_palette(list(wide.index))
        for b in wide.index:
            if b == "Westminster":
                continue
            r = wide.loc[b].tolist()
            figr.add_trace(go.Scatterpolar(r=r + [r[0]], theta=theta, name=b,
                                           line=dict(color=rpal[b], width=1.8), opacity=0.9))
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

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — CHILDREN & SCHOOLS
# Pupil numbers, the shift to independent schools, cross-border movement and SEN.
# Comparators: Westminster · 5 CIPFA neighbours · London · England.
# ══════════════════════════════════════════════════════════════════════════════
PHASES_MAIN = ["State-funded primary", "State-funded secondary", "State-funded nursery",
               "State-funded special school", "Independent school", "All phases"]

with tab5:
    st.subheader("Children & schools — a shrinking school-age population")
    st.markdown(
        "**Datasets:** DfE *School pupils and their characteristics*, DfE *Special educational "
        "needs in England*, and DfE *cross-border movement*. Pupil totals for **London** and "
        "**England** are taken from the SEN publication, which reports the total headcount for "
        "every phase nationally, regionally and for all 33 London boroughs — so Westminster can "
        "be read against its five CIPFA neighbours **and** the London and England averages.")

    if df_sen.empty and df_pupils.empty:
        st.info("No schools data found. Add `SEN_data-special-educational-needs-in-england.csv` "
                "and `data-school-pupils-and-their-characteristics.csv` to the `data` folder.")
    else:
        # ── headline metrics ──────────────────────────────────────────────────
        if not df_sen.empty:
            prim = sen_totals(df_sen, "State-funded primary")
            sec = sen_totals(df_sen, "State-funded secondary")
            yrs = sorted(prim["year"].unique())
            y0, y1 = yrs[0], yrs[-1]
            def _chg(d, geo):
                a = d[(d["geo"] == geo) & (d["year"] == y0)]["count"].sum()
                b = d[(d["geo"] == geo) & (d["year"] == y1)]["count"].sum()
                return (b / a - 1) * 100 if a else np.nan, b
            pw, pw_now = _chg(prim, "Westminster")
            sw, sw_now = _chg(sec, "Westminster")
            pl, _ = _chg(prim, LONDON)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"State-funded primary pupils ({_academic_year_label(str(y1)+str(y1+1)[2:])})",
                      f"{int(pw_now):,}" if pd.notna(pw_now) else "—",
                      delta=f"{pw:+.1f}% since {y0}/{str(y0+1)[2:]}" if pd.notna(pw) else None,
                      help="Westminster resident-funded primary headcount and its change since the baseline year.")
            m2.metric(f"State-funded secondary pupils",
                      f"{int(sw_now):,}" if pd.notna(sw_now) else "—",
                      delta=f"{sw:+.1f}% since {y0}/{str(y0+1)[2:]}" if pd.notna(sw) else None)
            m3.metric("Primary change — London", f"{pl:+.1f}%" if pd.notna(pl) else "—",
                      help="London-wide primary change over the same period, for context.")
            allt = sen_totals(df_sen, "All phases")
            ind = sen_totals(df_sen, "Independent school")
            iw = ind[(ind["geo"] == "Westminster") & (ind["year"] == y1)]["count"].sum()
            aw = allt[(allt["geo"] == "Westminster") & (allt["year"] == y1)]["count"].sum()
            m4.metric("Pupils in independent schools", f"{(iw/aw*100):.1f}%" if aw else "—",
                      help="Westminster share of all pupils attending independent schools — "
                           "roughly three times the London average.")

        st.divider()
        # ── 1 · % change vs baseline (PowerPoint slides 3 & 4) ────────────────
        st.markdown("### 1 · Change in pupil numbers against a baseline year")
        st.markdown(
            "This is the core measure in the child-population-decline analysis: each area's pupil "
            "headcount indexed to its own baseline, so areas of very different size can be compared "
            "on the same axis. Westminster's decline is far steeper than London's or England's.")
        if df_sen.empty:
            st.info("SEN/pupil totals file not found — needed for this chart.")
        else:
            c1, c2, c3 = st.columns(3)
            phase_sel = c1.selectbox("School phase", PHASES_MAIN, index=0, key="sch_phase")
            years_all = sorted(sen_totals(df_sen, phase_sel)["year"].unique())
            base_sel = c2.selectbox("Baseline year", years_all, index=0, key="sch_base",
                                    format_func=lambda y: f"{y}/{str(y+1)[2:]}")
            mode_sel = c3.selectbox("Measure", ["% change vs baseline", "Index (baseline = 100)",
                                                "Headcount"], key="sch_mode")
            geo_sel = st.multiselect("Areas to compare", COMPARATORS, default=COMPARATORS,
                                     key="sch_geos")
            d = sen_totals(df_sen, phase_sel)
            d = d[d["geo"].isin(geo_sel)]
            if d.empty:
                st.info("No data for that combination of phase and areas.")
            else:
                if mode_sel == "Headcount":
                    plot = d.rename(columns={"count": "value"})
                    ylab, fmt = "Pupil headcount", ":,"
                else:
                    plot = index_to_baseline(d, "count", base_sel,
                                             "pct_change" if mode_sel.startswith("%") else "index")
                    ylab = ("% change since baseline" if mode_sel.startswith("%")
                            else "Index (baseline = 100)")
                    fmt = ":.1f"
                chart_title(
                    f"{phase_sel} — {mode_sel.lower()}"
                    + (f" (baseline {base_sel}/{str(base_sel+1)[2:]})" if mode_sel != "Headcount" else ""),
                    "Westminster in strong colour · boroughs in their own muted colours · "
                    "London and England dashed")
                pal = borough_palette(sorted(plot["geo"].unique()))
                figp = go.Figure()
                for g in sorted(plot["geo"].unique(), key=lambda x: (x in AVERAGE_COLOURS, x)):
                    s = plot[plot["geo"] == g].sort_values("year")
                    figp.add_trace(go.Scatter(
                        x=s["year_label"], y=s["value"], mode="lines+markers", name=g,
                        line=dict(color=pal[g], **line_style(g)),
                        marker=dict(size=8 if g == "Westminster" else 5),
                        hovertemplate="<b>" + g + "</b><br>%{x}: %{y" + fmt + "}<extra></extra>"))
                if mode_sel != "Headcount":
                    figp.add_hline(y=0 if mode_sel.startswith("%") else 100,
                                   line_dash="dot", line_color="#BBBBBB")
                figp.update_xaxes(title="Academic year")
                figp.update_yaxes(title=ylab)
                show_chart(figp, "sch_change", "DfE pupil / SEN statistics")
                legend_hint("Click a borough in the legend to hide it, or double-click to isolate it.")

        st.divider()
        # ── 2 · Independent schools (slides 5-7) ─────────────────────────────
        st.markdown("### 2 · The independent-school share")
        st.markdown(
            "Westminster has one of the highest independent-school shares in the country. This "
            "matters for planning: a falling state-school roll is driven both by fewer children "
            "overall and by where those children are educated.")
        if df_sen.empty:
            st.info("Pupil totals file not found — needed for the independent-school share.")
        else:
            alltot = sen_totals(df_sen, "All phases").rename(columns={"count": "total"})
            indtot = sen_totals(df_sen, "Independent school").rename(columns={"count": "indep"})
            shr = alltot.merge(indtot, on=["year", "year_label", "geo"], how="left")
            shr["indep"] = shr["indep"].fillna(0)
            shr["share"] = np.where(shr["total"] > 0, shr["indep"] / shr["total"] * 100, np.nan)

            ca, cb = st.columns([3, 2])
            with ca:
                sel2 = st.multiselect("Areas", COMPARATORS, default=COMPARATORS, key="ind_geos")
                sh = shr[shr["geo"].isin(sel2)].dropna(subset=["share"])
                if not sh.empty:
                    chart_title("Share of pupils attending independent schools",
                                "% of all pupils in the area · Westminster in strong colour")
                    pal2 = borough_palette(sorted(sh["geo"].unique()))
                    figi = go.Figure()
                    for g in sorted(sh["geo"].unique(), key=lambda x: (x in AVERAGE_COLOURS, x)):
                        s = sh[sh["geo"] == g].sort_values("year")
                        figi.add_trace(go.Scatter(
                            x=s["year_label"], y=s["share"], mode="lines+markers", name=g,
                            line=dict(color=pal2[g], **line_style(g)),
                            marker=dict(size=8 if g == "Westminster" else 5),
                            hovertemplate="<b>" + g + "</b><br>%{x}: %{y:.1f}%<extra></extra>"))
                    figi.update_xaxes(title="Academic year")
                    figi.update_yaxes(title="% of pupils in independent schools", rangemode="tozero")
                    show_chart(figi, "indep_share_ts", "DfE pupil / SEN statistics")
            with cb:
                yr_i = st.selectbox("Year (ranking)", sorted(shr["year"].unique(), reverse=True),
                                    key="ind_year", format_func=lambda y: f"{y}/{str(y+1)[2:]}")
                rank = shr[(shr["year"] == yr_i) & (~shr["geo"].isin([ENGLAND, LONDON]))]
                rank = rank.dropna(subset=["share"]).sort_values("share")
                if not rank.empty:
                    chart_title(f"All London boroughs ranked, {yr_i}/{str(yr_i+1)[2:]}",
                                "% of pupils in independent schools · Westminster highlighted")
                    cols_r = np.where(rank["geo"] == "Westminster", FOCAL, CONTEXT_BAR)
                    figr = go.Figure(go.Bar(
                        x=rank["share"], y=rank["geo"], orientation="h", marker_color=cols_r,
                        hovertemplate="<b>%{y}</b><br>%{x:.1f}% independent<extra></extra>"))
                    figr.update_xaxes(title="% of pupils")
                    figr.update_yaxes(title="")
                    figr.update_layout(height=680)
                    show_chart(figr, "indep_rank", "DfE pupil / SEN statistics")

            # borough choropleth of independent share
            if borough_gj is not None:
                id_by_name = {}
                for ft in borough_gj["features"]:
                    p = ft["properties"]
                    nm = (p.get("name") or p.get("BoroughNa") or p.get("NAME") or
                          next((v for k, v in p.items() if isinstance(v, str) and "E09" not in v), ""))
                    id_by_name[_norm_la(nm)] = ft["id"]
                mm = shr[(shr["year"] == yr_i) & (~shr["geo"].isin([ENGLAND, LONDON]))].copy()
                mm["gid"] = mm["geo"].map(id_by_name)
                mm = mm.dropna(subset=["gid", "share"])
                if not mm.empty:
                    chart_title(f"Independent-school share across London, {yr_i}/{str(yr_i+1)[2:]}",
                                "Darker = a higher share of pupils educated privately")
                    figm = choropleth(borough_gj, mm["gid"], mm["share"], mm["geo"],
                                      "% independent", [[0, WCC["light_blue"]], [1, FOCAL]],
                                      fmt=":.1f", zoom=9, center={"lat": 51.50, "lon": -0.12},
                                      height=520)
                    show_chart(figm, "indep_map", "DfE pupil / SEN statistics")

        st.divider()
        # ── 3 · Cross-border movement ────────────────────────────────────────
        st.markdown("### 3 · Cross-border movement — where resident pupils go to school")
        st.markdown(
            "Resident pupils are not the same as pupils on local school rolls. This table splits "
            "each borough's **resident** maintained-school pupils into those educated **inside** the "
            "borough and those travelling **outside** it — the basis for Westminster's own "
            "independent-school estimate (LA resident headcount minus maintained-school residents).")
        if df_xborder.empty:
            st.info("Cross-border file not found — add "
                    "`cross_border_data_data-school-pupils-and-their-characteristics.csv` to `data`.")
        else:
            cx1, cx2 = st.columns(2)
            ph_x = cx1.selectbox("Phase", sorted(df_xborder["phase"].unique()), key="xb_phase")
            yr_x = cx2.selectbox("Academic year", sorted(df_xborder["year"].unique(), reverse=True),
                                 key="xb_year", format_func=lambda y: f"{y}/{str(y+1)[2:]}")
            xb = df_xborder[(df_xborder["phase"] == ph_x) & (df_xborder["year"] == yr_x)].copy()
            if not xb.empty:
                xb["pct_out"] = np.where(xb["resident"] > 0, xb["out_la"] / xb["resident"] * 100, np.nan)
                xb = xb.sort_values("pct_out")
                chart_title(f"Share of resident {ph_x.lower()} pupils educated outside their borough",
                            f"{yr_x}/{str(yr_x+1)[2:]} · maintained schools only · Westminster highlighted")
                colx = np.where(xb["la"] == "Westminster", FOCAL, CONTEXT_BAR)
                figx = go.Figure(go.Bar(
                    x=xb["pct_out"], y=xb["la"], orientation="h", marker_color=colx,
                    text=[f"{v:.0f}%" for v in xb["pct_out"]], textposition="outside",
                    customdata=np.stack([xb["resident"], xb["in_la"], xb["out_la"]], axis=-1),
                    hovertemplate="<b>%{y}</b><br>Resident pupils: %{customdata[0]:,.0f}"
                                  "<br>Schooled in borough: %{customdata[1]:,.0f}"
                                  "<br>Schooled outside: %{customdata[2]:,.0f}"
                                  "<br>= %{x:.1f}% travelling out<extra></extra>"))
                figx.update_xaxes(title="% of resident pupils schooled outside the borough",
                                  range=[0, float(np.nanmax(xb["pct_out"])) * 1.25])
                figx.update_yaxes(title="")
                figx.update_layout(height=330)
                show_chart(figx, "xborder_bar", "DfE cross-border movement")

            st.markdown("#### Independent-school estimate — calculated from source data")
            st.markdown(
                "This is **computed in the app**, not read from a spreadsheet:\n\n"
                "> independent = **LA resident headcount** (ONS mid-year estimates, ages 4–10 for "
                "primary and 11–16 for secondary) − **maintained-school resident headcount** "
                "(DfE cross-border movement)\n\n"
                "Each academic year is matched to the mid-year estimate at its **end** "
                "(2023/24 → mid-2024). `Independent_schools_calculations.xlsx` is used purely as a "
                "**reference to check the calculation against**, in the reconciliation table below.")
            if df_indep_calc.empty:
                st.info(
                    "The calculation needs both the cross-border file and a single-year-of-age "
                    "population source. Add `population_single_year_of_age.csv` (Nomis NM_2002_1) to "
                    "`data/` to compute this for all six boroughs; without it the app falls back to "
                    "the Westminster-only small-area MYE workbook.")
            else:
                src_used = ", ".join(sorted(set(df_indep_calc["source"].dropna())))
                st.caption(f"Population source in use: **{src_used}**")
                ic2 = df_indep_calc[df_indep_calc["phase"].str.lower() == ph_x.lower()]
                if not ic2.empty:
                    cc_a, cc_b = st.columns(2)
                    with cc_a:
                        chart_title(f"Estimated independent-school pupils — {ph_x.lower()}",
                                    "Resident children minus those in maintained schools")
                        pal_i = borough_palette(sorted(ic2["la"].unique()))
                        fig_i1 = go.Figure()
                        for g in sorted(ic2["la"].unique()):
                            sg = ic2[ic2["la"] == g].sort_values("year")
                            fig_i1.add_trace(go.Scatter(
                                x=sg["year_label"], y=sg["indep_estimate"], mode="lines+markers",
                                name=g, line=dict(color=pal_i[g], **line_style(g)),
                                hovertemplate="<b>" + g + "</b><br>%{x}: %{y:,.0f} pupils<extra></extra>"))
                        fig_i1.update_xaxes(title="Academic year")
                        fig_i1.update_yaxes(title="Estimated independent-school pupils")
                        show_chart(fig_i1, "indep_calc_n", "Calculated: ONS MYE − DfE cross-border")
                    with cc_b:
                        chart_title(f"Estimated % in independent schools — {ph_x.lower()}",
                                    "Share of the borough's resident children of that age")
                        fig_i2 = go.Figure()
                        for g in sorted(ic2["la"].unique()):
                            sg = ic2[ic2["la"] == g].sort_values("year")
                            fig_i2.add_trace(go.Scatter(
                                x=sg["year_label"], y=sg["pct_independent"], mode="lines+markers",
                                name=g, line=dict(color=pal_i[g], **line_style(g)),
                                hovertemplate="<b>" + g + "</b><br>%{x}: %{y:.1f}%<extra></extra>"))
                        fig_i2.update_xaxes(title="Academic year")
                        fig_i2.update_yaxes(title="% of resident children")
                        show_chart(fig_i2, "indep_calc_pct", "Calculated: ONS MYE − DfE cross-border")

                # ── reconciliation against the reference spreadsheet
                if not df_indep.empty:
                    st.markdown("**Reconciliation against `Independent_schools_calculations.xlsx`**")
                    ref = df_indep[["year", "la", "phase", "la_resident", "indep_estimate"]].rename(
                        columns={"la_resident": "ref_resident", "indep_estimate": "ref_indep"})
                    rec = df_indep_calc.merge(ref, on=["year", "la", "phase"], how="inner")
                    if rec.empty:
                        st.info("No overlapping rows between the calculation and the reference table.")
                    else:
                        rec["ref_pct"] = np.where(rec["ref_resident"] > 0,
                                                  rec["ref_indep"] / rec["ref_resident"] * 100, np.nan)
                        rec["diff_pp"] = rec["pct_independent"] - rec["ref_pct"]
                        show = rec[["year_label", "la", "phase", "la_resident", "ref_resident",
                                    "pct_independent", "ref_pct", "diff_pp"]].copy()
                        show.columns = ["Academic year", "LA", "Phase", "MYE resident (computed)",
                                        "Resident (reference)", "% independent (computed)",
                                        "% independent (reference)", "Difference (pp)"]
                        st.dataframe(
                            show.style.format({
                                "MYE resident (computed)": "{:,.0f}", "Resident (reference)": "{:,.0f}",
                                "% independent (computed)": "{:.1f}%", "% independent (reference)": "{:.1f}%",
                                "Difference (pp)": "{:+.1f}"}),
                            use_container_width=True, hide_index=True)
                        worst = rec["diff_pp"].abs().max()
                        flagged = rec[rec["diff_pp"].abs() > 1.0]
                        if worst <= 1.0:
                            st.success(
                                f"✅ **Calculation validated.** Every overlapping row matches the "
                                f"reference table to within {worst:.1f} percentage points — small "
                                "differences are expected because the reference used the LA-level MYE "
                                "series while the app uses the LSOA-rebased small-area estimates.")
                        else:
                            rows = ", ".join(f"{r.la} {r.phase} {r.year_label}"
                                             for r in flagged.itertuples())
                            st.warning(
                                f"⚠️ **Check these rows:** {rows} differ from the reference by more "
                                f"than 1 percentage point (largest gap {worst:.1f}pp). Note that in "
                                "the reference spreadsheet the final year (2023/24) was **not** "
                                "calculated — its `MYE 4-10`/`MYE 11-16` cells are blank and its "
                                "percentage was carried over unchanged from 2022/23, with the resident "
                                "headcount back-derived from it. The app's figure for that year is a "
                                "genuine calculation, so a gap there is expected and the app's value "
                                "is the more reliable one.")

        st.divider()
        # ── 4 · SEN (slides 18-20) ───────────────────────────────────────────
        st.markdown("### 4 · Special educational needs")
        st.markdown(
            "Two things move independently here: the **number** of SEN pupils (which follows the "
            "overall roll) and the **rate** of SEN (the share of pupils with an EHC plan or on SEN "
            "support). Westminster's SEN headcount has fallen with its shrinking roll even as the "
            "SEN *rate* has risen.")
        if df_sen.empty:
            st.info("SEN file not found — add `SEN_data-special-educational-needs-in-england.csv`.")
        else:
            s1, s2, s3 = st.columns(3)
            ph_s = s1.selectbox("Phase", [p for p in PHASES_MAIN if p != "Independent school"],
                                key="sen_phase")
            meas_s = s2.selectbox("Measure", ["SEN rate (% of pupils)",
                                              "SEN headcount (indexed to baseline = 100)",
                                              "SEN headcount"], key="sen_meas")
            prov_s = s3.selectbox("SEN provision", ["All SEN", "Education, health and care plan",
                                                    "SEN support / SEN without an EHC plan"],
                                  key="sen_prov")
            base = df_sen[df_sen["phase"] == ph_s] if ph_s != "All phases" else df_sen
            tot = base[base["provision"] == "Total"].groupby(
                ["year", "year_label", "geo"], as_index=False)["count"].sum().rename(columns={"count": "total"})
            if prov_s == "All SEN":
                sen_c = base[base["provision"] != "Total"]
            else:
                sen_c = base[base["provision"] == prov_s]
            sen_c = sen_c.groupby(["year", "year_label", "geo"], as_index=False)["count"].sum().rename(
                columns={"count": "sen"})
            mg = tot.merge(sen_c, on=["year", "year_label", "geo"], how="left")
            mg["sen"] = mg["sen"].fillna(0)
            mg["rate"] = np.where(mg["total"] > 0, mg["sen"] / mg["total"] * 100, np.nan)
            sel_s = st.multiselect("Areas", COMPARATORS, default=COMPARATORS, key="sen_geos")
            mg = mg[mg["geo"].isin(sel_s)]
            if mg.empty:
                st.info("No SEN data for that combination.")
            else:
                if meas_s.startswith("SEN rate"):
                    mg["value"] = mg["rate"]; ylab = "% of pupils with SEN"; fmt = ":.1f"
                elif meas_s.startswith("SEN headcount (indexed"):
                    ix = index_to_baseline(mg.rename(columns={"sen": "count"}), "count",
                                           mg["year"].min(), "index")
                    mg = ix; ylab = "Indexed SEN headcount (baseline = 100)"; fmt = ":.1f"
                else:
                    mg["value"] = mg["sen"]; ylab = "SEN pupils (headcount)"; fmt = ":,"
                chart_title(f"{meas_s} — {ph_s.lower()}, {prov_s.lower()}",
                            "Westminster in strong colour · London and England dashed")
                pal3 = borough_palette(sorted(mg["geo"].unique()))
                figs = go.Figure()
                for g in sorted(mg["geo"].unique(), key=lambda x: (x in AVERAGE_COLOURS, x)):
                    s = mg[mg["geo"] == g].sort_values("year")
                    figs.add_trace(go.Scatter(
                        x=s["year_label"], y=s["value"], mode="lines+markers", name=g,
                        line=dict(color=pal3[g], **line_style(g)),
                        marker=dict(size=8 if g == "Westminster" else 5),
                        hovertemplate="<b>" + g + "</b><br>%{x}: %{y" + fmt + "}<extra></extra>"))
                if meas_s.startswith("SEN headcount (indexed"):
                    figs.add_hline(y=100, line_dash="dot", line_color="#BBBBBB")
                figs.update_xaxes(title="Academic year")
                figs.update_yaxes(title=ylab)
                show_chart(figs, "sen_trend", "DfE special educational needs in England")
                legend_hint()

            # SEN provision split for Westminster (EHCP vs SEN support)
            wsen = df_sen[(df_sen["geo"] == "Westminster") & (df_sen["provision"] != "Total")]
            if ph_s != "All phases":
                wsen = wsen[wsen["phase"] == ph_s]
            wsen = wsen.groupby(["year_label", "provision"], as_index=False)["count"].sum()
            if not wsen.empty:
                chart_title("Westminster's SEN pupils by type of provision",
                            "EHC plans vs SEN support · stacked headcount")
                figsp = px.bar(wsen, x="year_label", y="count", color="provision",
                               color_discrete_map={"Education, health and care plan": FOCAL,
                                                   "SEN support / SEN without an EHC plan": "#8598CE"})
                figsp.update_xaxes(title="Academic year")
                figsp.update_yaxes(title="SEN pupils")
                show_chart(figsp, "sen_split", "DfE special educational needs in England")

        st.markdown("#### SEN by ethnicity")
        st.markdown(
            "Ethnicity is where Westminster diverges most sharply from London: London has seen SEN "
            "numbers rise across nearly every ethnic group, while Westminster's changes are mixed "
            "and, for several groups, strongly negative. Because the overall roll is shrinking, read "
            "the **rate** alongside the **count** — a falling count can sit with a rising rate.")
        if df_sen_eth.empty:
            st.info("SEN-by-ethnicity file not found - add `sen_fsm_eth_lang_new_.csv` to `data/`.")
        else:
            if EXCLUDE_CITY_LONDON:
                df_sen_eth = df_sen_eth[df_sen_eth["geo"] != "City of London"]
            e1, e2, e3 = st.columns(3)
            ph_opts = ["All phases"] + sorted(df_sen_eth["phase"].unique())
            ph_e = e1.selectbox("Phase", ph_opts,
                                index=ph_opts.index("State-funded primary")
                                if "State-funded primary" in ph_opts else 0, key="eth_phase")
            prov_opts = sorted(df_sen_eth["provision"].unique())
            prov_e = e2.selectbox("SEN provision", prov_opts, key="eth_prov")
            meas_e = e3.selectbox("Measure", ["% change over the period", "SEN pupils (count)"],
                                  key="eth_meas")
            de = df_sen_eth.copy()
            if ph_e != "All phases":
                de = de[de["phase"] == ph_e]
            de = de[de["provision"] == prov_e]
            de = de.groupby(["year", "year_label", "geo", "ethnicity"], as_index=False)["count"].sum()
            geo_e = st.multiselect("Areas to compare", sorted(de["geo"].unique()),
                                   default=[g for g in ["Westminster", LONDON, ENGLAND]
                                            if g in set(de["geo"])] or sorted(de["geo"].unique())[:2],
                                   key="eth_geos")
            de = de[de["geo"].isin(geo_e)]
            if de.empty:
                st.info("No SEN-by-ethnicity data for that combination.")
            else:
                yrs_e = sorted(de["year"].unique())
                if meas_e.startswith("%") and len(yrs_e) > 1:
                    ya, yb = st.select_slider("Period", options=yrs_e, value=(yrs_e[0], yrs_e[-1]),
                                              key="eth_period",
                                              format_func=lambda y: f"{int(y)}/{str(int(y)+1)[2:]}")
                    a = de[de["year"] == ya].set_index(["geo", "ethnicity"])["count"]
                    b = de[de["year"] == yb].set_index(["geo", "ethnicity"])["count"]
                    ch = pd.concat([a.rename("start"), b.rename("end")], axis=1).dropna()
                    ch = ch[ch["start"] > 0].reset_index()
                    ch["pct"] = (ch["end"] / ch["start"] - 1) * 100
                    if ch.empty:
                        st.info("Not enough overlapping data to compute change for that period.")
                    else:
                        order = (ch[ch["geo"] == "Westminster"].sort_values("pct")["ethnicity"].tolist()
                                 if "Westminster" in set(ch["geo"]) else
                                 ch.groupby("ethnicity")["pct"].mean().sort_values().index.tolist())
                        chart_title(
                            f"Change in SEN pupils by ethnicity, {int(ya)}/{str(int(ya)+1)[2:]} to "
                            f"{int(yb)}/{str(int(yb)+1)[2:]} — {ph_e.lower()}",
                            "% change in the number of SEN pupils in each ethnic group · "
                            "bars right of zero grew, bars left of zero shrank")
                        pal_e = borough_palette(sorted(ch["geo"].unique()))
                        fige2 = go.Figure()
                        for g in sorted(ch["geo"].unique(), key=lambda x: (x in AVERAGE_COLOURS, x)):
                            sg = ch[ch["geo"] == g].set_index("ethnicity").reindex(order).reset_index()
                            fige2.add_trace(go.Bar(
                                y=sg["ethnicity"], x=sg["pct"], orientation="h", name=g,
                                marker_color=pal_e[g],
                                hovertemplate="<b>%{y}</b><br>" + g + ": %{x:+.1f}%<extra></extra>"))
                        fige2.add_vline(x=0, line_color="#888888")
                        fige2.update_xaxes(title="% change in SEN pupils")
                        fige2.update_yaxes(title="")
                        fige2.update_layout(barmode="group", height=max(420, 34 * len(order)))
                        show_chart(fige2, "sen_eth_change", "DfE SEN in England, by ethnicity")
                        legend_hint()
                        if "Westminster" in set(ch["geo"]):
                            wch = ch[ch["geo"] == "Westminster"]
                            down = wch[wch["pct"] < 0].sort_values("pct")
                            if len(down):
                                names = ", ".join(f"{r.ethnicity} ({r.pct:+.0f}%)"
                                                  for r in down.head(3).itertuples())
                                st.info(f"**Largest falls in Westminster:** {names}. Compare each "
                                        "against the London bar beside it — where London rose and "
                                        "Westminster fell, the divergence is specific to the borough "
                                        "rather than a London-wide trend.")
                else:
                    chart_title(f"SEN pupils by ethnicity over time — {ph_e.lower()}",
                                "Counts by ethnic group · use the legend to isolate a group")
                    tot_e = de.groupby(["year_label", "ethnicity"], as_index=False)["count"].sum()
                    fige3 = px.line(tot_e.sort_values("year_label"), x="year_label", y="count",
                                    color="ethnicity", markers=True)
                    fige3.update_xaxes(title="Academic year")
                    fige3.update_yaxes(title="SEN pupils")
                    show_chart(fige3, "sen_eth_ts", "DfE SEN in England, by ethnicity")
        if not df_fsm_lang.empty:
            st.markdown("**FSM eligibility and first language, alongside the SEN picture**")
            fl = df_fsm_lang.copy()
            if EXCLUDE_CITY_LONDON:
                fl = fl[fl["geo"] != "City of London"]
            fl = fl[(fl["provision"] == "Total") &
                    (fl["metric"].isin(["FSM eligible", "Other first language"]))]
            f1, f2 = st.columns(2)
            metric_f = f1.selectbox("Metric", ["FSM eligible", "Other first language"], key="fl_metric")
            geo_f = f2.multiselect("Areas", sorted(fl["geo"].unique()),
                                   default=[g for g in ["Westminster", "London"] if g in set(fl["geo"])],
                                   key="fl_geos")
            dff = fl[(fl["metric"] == metric_f) & (fl["geo"].isin(geo_f))]
            dff = dff.groupby(["year", "year_label", "geo"], as_index=False)["count"].sum()
            if not dff.empty:
                chart_title(f"Pupils recorded as '{metric_f.lower()}' over time",
                            "Westminster in strong colour, London dashed")
                palf = borough_palette(sorted(dff["geo"].unique()))
                figf = go.Figure()
                for g in sorted(dff["geo"].unique(), key=lambda x: (x in AVERAGE_COLOURS, x)):
                    s = dff[dff["geo"] == g].sort_values("year")
                    figf.add_trace(go.Scatter(
                        x=s["year_label"], y=s["count"], mode="lines+markers", name=g,
                        line=dict(color=palf.get(g, CONTEXT_LINE), **line_style(g)),
                        hovertemplate="<b>" + g + "</b><br>%{x}: %{y:,.0f}<extra></extra>"))
                figf.update_xaxes(title="Academic year")
                figf.update_yaxes(title="Pupils")
                show_chart(figf, "fsm_lang_trend", "DfE SEN/FSM/language by ethnicity, London")
        source_line("Pupil totals, independent-school shares and SEN counts: DfE Explore Education "
                    "Statistics (School pupils and their characteristics; Special educational needs "
                    "in England; cross-border movement). London and England totals are the regional "
                    "and national rows of the SEN publication.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — CHILDCARE COSTS
# DfE Childcare and early years provider survey — hourly fees across London LAs.
# ══════════════════════════════════════════════════════════════════════════════
with tab6:
    st.subheader("Childcare costs across London")
    st.markdown(
        "**Dataset:** DfE *Childcare and early years provider survey* — the hourly fee charged by "
        "providers, for **2-year-olds** and **3-and-4-year-olds**, in every London borough. "
        "Childcare cost is one of the pressures behind families leaving high-cost inner London, so "
        "it sits alongside the births and migration evidence in the next tab. "
        "The survey publishes London local authorities only, so England-wide comparison is not "
        "available here; the **London average** shown is the mean across boroughs.")

    if df_ccosts.empty:
        st.info("Childcare cost file not found — add "
                "`costs_data-childcare-and-early-years-provider-survey.csv` to the `data` folder.")
    else:
        cc1, cc2, cc3 = st.columns(3)
        age_c = cc1.selectbox("Child age", sorted(df_ccosts["child_age"].unique()), key="cc_age")
        yr_c = cc2.selectbox("Year", sorted(df_ccosts["year"].dropna().unique(), reverse=True),
                             key="cc_year")
        meas_c = cc3.selectbox("Measure", ["Median hourly fee", "Mean hourly fee"], key="cc_meas")
        vcol = "median_fee" if meas_c.startswith("Median") else "mean_fee"

        d = df_ccosts[(df_ccosts["child_age"] == age_c) & (df_ccosts["year"] == yr_c)].copy()
        d = d.dropna(subset=[vcol])
        if d.empty:
            st.info("No childcare data for that combination.")
        else:
            ldn_avg = d[vcol].mean()
            wrow = d[d["la"] == "Westminster"]
            wval = float(wrow[vcol].iloc[0]) if len(wrow) else np.nan
            rank = int((d[vcol] > wval).sum() + 1) if pd.notna(wval) else None

            k1, k2, k3, k4 = st.columns(4)
            k1.metric(f"Westminster — {meas_c.lower()}",
                      f"£{wval:.2f}" if pd.notna(wval) else "—",
                      delta=f"{(wval-ldn_avg):+.2f} vs London avg" if pd.notna(wval) else None,
                      delta_color="inverse",
                      help="Hourly childcare fee for the selected age group. Red = more expensive "
                           "than the London average.")
            k2.metric("London average", f"£{ldn_avg:.2f}")
            k3.metric("Westminster rank in London",
                      f"{rank} of {len(d)}" if rank else "—",
                      help="1 = most expensive borough in London.")
            prev = df_ccosts[(df_ccosts["child_age"] == age_c) &
                             (df_ccosts["year"] == yr_c - 1) & (df_ccosts["la"] == "Westminster")]
            if len(prev) and pd.notna(wval):
                pv = float(prev[vcol].iloc[0])
                k4.metric(f"Change since {yr_c-1}", f"{(wval/pv-1)*100:+.1f}%",
                          delta_color="inverse")
            else:
                k4.metric(f"Change since {yr_c-1}", "—")

            # ── ranked bar across all London boroughs
            dr = d.sort_values(vcol)
            chart_title(f"{meas_c} for {age_c}, {yr_c}",
                        "Every London borough ranked · Westminster highlighted · "
                        "dashed line = London average")
            colc = np.where(dr["la"] == "Westminster", FOCAL, CONTEXT_BAR)
            figc = go.Figure(go.Bar(
                x=dr[vcol], y=dr["la"], orientation="h", marker_color=colc,
                text=[f"£{v:.2f}" for v in dr[vcol]], textposition="outside",
                hovertemplate="<b>%{y}</b><br>" + meas_c + ": £%{x:.2f}<extra></extra>"))
            figc.add_vline(x=ldn_avg, line_dash="dash", line_color=AVERAGE_COLOURS[LONDON])
            figc.update_xaxes(title=f"{meas_c} (£)", range=[0, float(dr[vcol].max()) * 1.18])
            figc.update_yaxes(title="")
            figc.update_layout(height=700)
            show_chart(figc, "cc_rank", "DfE childcare and early years provider survey")

            # ── London choropleth
            if borough_gj is not None:
                id_by_code = {}
                for ft in borough_gj["features"]:
                    p = ft["properties"]
                    code = next((v for v in p.values()
                                 if isinstance(v, str) and v.startswith("E09")), None)
                    if code:
                        id_by_code[code] = ft["id"]
                dm = d.copy()
                dm["gid"] = dm["la_code"].map(id_by_code)
                if dm["gid"].isna().all():
                    id_by_name = {}
                    for ft in borough_gj["features"]:
                        p = ft["properties"]
                        nm = (p.get("name") or p.get("BoroughNa") or p.get("NAME") or
                              next((v for k, v in p.items() if isinstance(v, str) and "E09" not in v), ""))
                        id_by_name[_norm_la(nm)] = ft["id"]
                    dm["gid"] = dm["la"].map(id_by_name)
                dm = dm.dropna(subset=["gid"])
                if not dm.empty:
                    chart_title(f"Where childcare costs most — {age_c}, {yr_c}",
                                "Darker = more expensive per hour · hover for the borough fee")
                    figmm = choropleth(borough_gj, dm["gid"], dm[vcol], dm["la"],
                                       f"{meas_c} (£)", [[0, WCC["light_blue"]], [1, FOCAL]],
                                       fmt=":.2f", zoom=9, center={"lat": 51.50, "lon": -0.12},
                                       height=520)
                    show_chart(figmm, "cc_map", "DfE childcare and early years provider survey")

        # ── age-group comparison + year-on-year change
        st.divider()
        st.markdown("### Comparing age groups and years")
        cmp_geo = st.multiselect(
            "Areas", sorted(df_ccosts["la"].unique()),
            default=[g for g in ["Westminster"] + [n for n in NEIGHBOURS if n != "Westminster"]
                     if g in set(df_ccosts["la"])], key="cc_geos")
        dd = df_ccosts[df_ccosts["la"].isin(cmp_geo)].dropna(subset=[vcol])
        if not dd.empty:
            chart_title(f"{meas_c} by age group and year",
                        "Grouped by borough · 2-year-olds cost more per hour than 3-and-4-year-olds")
            dd = dd.copy()
            dd["grp"] = dd["child_age"] + " · " + dd["year"].astype(str)
            figg = px.bar(dd.sort_values("la"), x="la", y=vcol, color="grp", barmode="group",
                          color_discrete_sequence=[FOCAL, "#5FA8A0", "#D0A44C", "#C57B8A"])
            figg.update_xaxes(title="")
            figg.update_yaxes(title=f"{meas_c} (£)")
            show_chart(figg, "cc_group", "DfE childcare and early years provider survey")

            # affordability framing
            st.info("**Reading these figures:** hourly fees are only half the affordability picture — "
                    "the other half is local earnings. Westminster combines high fees with high median "
                    "pay, so its *affordability ratio* is better than boroughs with similar fees but "
                    "lower wages. Where a fee-to-earnings ratio is needed, pair this with ONS "
                    "median hourly pay by borough.")

        source_line("Childcare fees: DfE Childcare and early years provider survey (London local "
                    "authorities). Fees are the hourly amount charged by providers, not the amount "
                    "paid by parents after free-entitlement hours and subsidies.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — BIRTHS, MIGRATION & CHILD POPULATION DECLINE
# The demographic drivers: fewer births, net outflow of families, and the
# resulting decline in the school-age cohort — with a simple forward projection.
# ══════════════════════════════════════════════════════════════════════════════
with tab7:
    st.subheader("Why the child population is falling")
    st.markdown(
        "Three forces drive Westminster's shrinking child population: **fewer births**, "
        "**net domestic out-migration of families**, and the way those two feed through into "
        "each successive school cohort. This tab brings them together and projects the trend "
        "forward.")

    # ── headline metrics ──────────────────────────────────────────────────────
    mk1, mk2, mk3, mk4 = st.columns(4)
    if not df_mye_la.empty:
        w = df_mye_la[(df_mye_la["area"] == "Westminster") & (df_mye_la["gender"] == "Total")]
        w = w.groupby("year", as_index=False)["population"].sum().sort_values("year")
        if len(w) > 1:
            latest, first = w.iloc[-1], w.iloc[0]
            peak = w.loc[w["population"].idxmax()]
            mk1.metric(f"Children 0–19 ({int(latest['year'])})", f"{int(latest['population']):,}",
                       delta=f"{(latest['population']/peak['population']-1)*100:+.1f}% vs {int(peak['year'])} peak",
                       delta_color="inverse",
                       help="ONS mid-year estimates, Westminster residents aged 0–19.")
            mk2.metric("Peak child population", f"{int(peak['population']):,} ({int(peak['year'])})")
    if not df_migr.empty:
        wm = df_migr[(df_migr["area_code"] == "E09000033") & (df_migr["age_band"] == "0-15")]
        if not wm.empty:
            net = float(wm["net"].iloc[0])
            mk3.metric("Net internal migration, 0–15 (2024)", f"{net:,.0f}",
                       delta="net outflow" if net < 0 else "net inflow", delta_color="inverse",
                       help="Children moving into Westminster from elsewhere in the UK minus those "
                            "moving out. Excludes international migration.")
    if not df_sen.empty:
        pr = sen_totals(df_sen, "State-funded primary")
        pw = pr[pr["geo"] == "Westminster"].sort_values("year")
        if len(pw) > 1:
            mk4.metric("State-funded primary roll",
                       f"{int(pw['count'].iloc[-1]):,}",
                       delta=f"{(pw['count'].iloc[-1]/pw['count'].iloc[0]-1)*100:+.1f}% since {int(pw['year'].iloc[0])}",
                       delta_color="inverse")

    st.divider()
    # ── 1 · Births ────────────────────────────────────────────────────────────
    st.markdown("### 1 · Births")
    st.markdown(
        "Births are the leading indicator: today's births set the size of the reception cohort in "
        "four years and the secondary cohort in eleven. Westminster's births have fallen sharply "
        "since the mid-2010s.")
    if df_births.empty:
        st.info(
            "**Births data not loaded.** The app looks for `births_by_age_of_mother.csv` in the "
            "`data` folder and, if it isn't there, calls the Nomis API (dataset NM_205_1) directly. "
            "Add the CSV to `data/` for a fast, offline-safe load — the API call is skipped whenever "
            "the file is present. Once loaded, this section shows total births by borough over time "
            "and the age-of-mother breakdown.")
    else:
        b1, b2 = st.columns(2)
        las_b = sorted(df_births["la"].unique())
        sel_b = b1.multiselect("Boroughs", las_b,
                               default=[l for l in las_b if l in set(COMPARATORS)] or las_b,
                               key="b_geos")
        ages_b = sorted(df_births["age_of_mother"].unique())
        age_b = b2.selectbox("Age of mother", ["All ages"] + [a for a in ages_b if a != "All ages"],
                             key="b_age")
        db = df_births[df_births["la"].isin(sel_b)]
        db = (db if age_b == "All ages" else db[db["age_of_mother"] == age_b])
        db = db.groupby(["year", "la"], as_index=False)["births"].sum()
        yr_min, yr_max = int(db["year"].min()), int(db["year"].max())
        rng_b = st.slider("Years", yr_min, yr_max, (yr_min, yr_max), key="b_years")
        db = db[(db["year"] >= rng_b[0]) & (db["year"] <= rng_b[1])]
        if not db.empty:
            chart_title(f"Live births by borough — {age_b.lower()}",
                        "Westminster in strong colour · comparators muted")
            palb = borough_palette(sorted(db["la"].unique()))
            figb = go.Figure()
            for g in sorted(db["la"].unique()):
                s = db[db["la"] == g].sort_values("year")
                figb.add_trace(go.Scatter(
                    x=s["year"], y=s["births"], mode="lines+markers", name=g,
                    line=dict(color=palb[g], **line_style(g)),
                    marker=dict(size=8 if g == "Westminster" else 5),
                    hovertemplate="<b>" + g + "</b><br>%{x}: %{y:,.0f} births<extra></extra>"))
            figb.update_xaxes(title="Year")
            figb.update_yaxes(title="Live births", rangemode="tozero")
            show_chart(figb, "births_ts", "ONS births via Nomis")
            legend_hint()

            # indexed view so boroughs of different size are comparable
            ib = db.rename(columns={"la": "geo", "births": "count"})
            ib = index_to_baseline(ib, "count", int(db["year"].min()), "index")
            if not ib.empty:
                chart_title("Births indexed to the first year shown (= 100)",
                            "Removes the size difference between boroughs to compare the rate of decline")
                figbi = go.Figure()
                for g in sorted(ib["geo"].unique()):
                    s = ib[ib["geo"] == g].sort_values("year")
                    figbi.add_trace(go.Scatter(
                        x=s["year"], y=s["value"], mode="lines", name=g,
                        line=dict(color=palb.get(g, CONTEXT_LINE), **line_style(g)),
                        hovertemplate="<b>" + g + "</b><br>%{x}: %{y:.1f}<extra></extra>"))
                figbi.add_hline(y=100, line_dash="dot", line_color="#BBBBBB")
                figbi.update_xaxes(title="Year"); figbi.update_yaxes(title="Index (first year = 100)")
                show_chart(figbi, "births_index", "ONS births via Nomis")

    # ── fertility rates (CBR · GFR · TFR) ─────────────────────────────────────
    st.markdown("**Fertility rates — CBR, GFR and TFR**")
    st.markdown(
        "Birth *counts* fall when there are fewer women of childbearing age, even if each woman has "
        "the same number of children. Fertility **rates** strip that out: the **GFR** is births per "
        "1,000 women aged 15–44, and the **TFR** is the number of children a woman would have across "
        "her lifetime at current age-specific rates. Westminster's TFR has been below the London and "
        "England averages throughout.")
    _fert_all = pd.concat(
        [d for d in [df_fert, derive_fertility_rates(df_births, df_fpop)] if d is not None and not d.empty],
        ignore_index=True) if (not df_fert.empty or (not df_births.empty and not df_fpop.empty)) \
        else pd.DataFrame(columns=["year", "area", "measure", "value"])
    if _fert_all.empty:
        st.info(
            "**Fertility rates not loaded.** The app reads Nomis dataset **NM_207_1** (total "
            "fertility rate, general fertility rate, crude birth rate and the age-specific rates) "
            "for all London boroughs plus London and England. It looks for "
            "`fertility_rates.csv` in `data/` first and calls the Nomis API if that is absent — so "
            "this message means both routes were unavailable (most likely no network access from "
            "the server). Saving the API response to `data/fertility_rates.csv` fixes it "
            "permanently. As an alternative the app can derive the GFR and TFR itself if you add "
            "`female_population_by_age.csv` alongside the births extract.")
    else:
        f1, f2 = st.columns(2)
        meas_f = f1.selectbox("Rate", order_fertility_measures(_fert_all["measure"].unique()),
                              key="fert_meas")
        areas_f = sorted(_fert_all["area"].unique())
        sel_f = f2.multiselect("Areas", areas_f,
                               default=[a for a in areas_f if a in set(COMPARATORS)] or areas_f,
                               key="fert_geos")
        df_f = _fert_all[(_fert_all["measure"] == meas_f) & (_fert_all["area"].isin(sel_f))]
        if df_f.empty:
            st.info("No values for that rate and area selection.")
        else:
            chart_title(f"{meas_f} over time",
                        "Westminster in strong colour · London and England dashed")
            pal_f = borough_palette(sorted(df_f["area"].unique()))
            fig_f = go.Figure()
            for g in sorted(df_f["area"].unique(), key=lambda x: (x in AVERAGE_COLOURS, x)):
                sg = df_f[df_f["area"] == g].sort_values("year")
                fig_f.add_trace(go.Scatter(
                    x=sg["year"], y=sg["value"], mode="lines+markers", name=g,
                    line=dict(color=pal_f[g], **line_style(g)),
                    marker=dict(size=8 if g == "Westminster" else 5),
                    hovertemplate="<b>" + g + "</b><br>%{x}: %{y:.2f}<extra></extra>"))
            if "fertility rate (TFR)" in meas_f:
                fig_f.add_hline(y=2.08, line_dash="dot", line_color="#BBBBBB",
                                annotation_text="replacement level (2.08)",
                                annotation_position="top left",
                                annotation_font=dict(size=10, color="#777777"))
            fig_f.update_xaxes(title="Year")
            fig_f.update_yaxes(title=meas_f)
            show_chart(fig_f, "fertility_rates", "ONS births and population estimates")
            legend_hint()

    # ── births by LSOA map ────────────────────────────────────────────────────
    st.markdown("**Births by Westminster LSOA, over time**")
    if df_births_lsoa.empty:
        st.info("LSOA-level births not loaded — add `births_by_lsoa.csv` (Nomis dataset NM_206_1) "
                "to the `data` folder to enable the small-area births map and its year slider.")
    elif lsoa_gj is None:
        st.info("LSOA boundary file not found — needed to draw the births map.")
    else:
        yrs_l = sorted(df_births_lsoa["year"].unique())
        yr_l = st.select_slider("Year", options=yrs_l, value=yrs_l[-1], key="bl_year")
        dl = df_births_lsoa[df_births_lsoa["year"] == yr_l].groupby(
            ["LSOA_CODE", "LSOA_NAME"], as_index=False)["births"].sum()
        dl = add_ward(dl)
        chart_title(f"Live births by LSOA, {int(yr_l)}",
                    "Darker = more births · hover for the ward · move the slider to see change over time")
        figbl = choropleth(lsoa_gj, dl["LSOA_CODE"], dl["births"], dl["LSOA_NAME"],
                           "Births", [[0, WCC["light_blue"]], [1, FOCAL]],
                           wards=dl["Ward"].tolist(), fmt=":,")
        show_chart(figbl, "births_lsoa_map", "ONS births by LSOA via Nomis")

    st.divider()
    # ── 2 · Domestic migration ───────────────────────────────────────────────
    st.markdown("### 2 · Domestic (internal) migration")
    st.markdown(
        "Internal migration is movement **within the UK**. Westminster loses children in every "
        "primary and secondary age band but gains sharply at 15–19, when students and young adults "
        "move in — so a single 'all ages' figure hides what is happening to families.")
    if df_migr.empty:
        st.info("Internal migration data not found. Add the small pre-aggregated file "
                "`internal_migration_children_2024.csv` to `data/` (recommended), or the full ONS "
                "workbook `detailedestimates2024on2023las.xlsx` — the app will aggregate it, but "
                "that takes several minutes on first load.")
    else:
        g1, g2 = st.columns(2)
        band_m = g1.selectbox("Age band", ["0-4", "5-9", "10-14", "15-19", "0-15", "All ages"],
                              index=4, key="mig_band")
        meas_m = g2.selectbox("Measure", ["Net migration", "Inflow", "Outflow"], key="mig_meas")
        mcol = {"Net migration": "net", "Inflow": "inflow", "Outflow": "outflow"}[meas_m]
        dm2 = df_migr[df_migr["age_band"] == band_m].copy()

        # Westminster's own age profile — the key insight
        wprof = df_migr[(df_migr["area_code"] == "E09000033") &
                        (df_migr["age_band"].isin(["0-4", "5-9", "10-14", "15-19"]))]
        if not wprof.empty:
            order = ["0-4", "5-9", "10-14", "15-19"]
            wprof = wprof.set_index("age_band").reindex(order).reset_index()
            chart_title("Westminster loses children but gains young adults (2024)",
                        "Net internal migration by age band · negative = more leaving than arriving")
            colsw = np.where(wprof["net"] < 0, FOCAL, "#94B36A")
            figw = go.Figure(go.Bar(
                x=wprof["age_band"], y=wprof["net"], marker_color=colsw,
                text=[f"{v:+,.0f}" for v in wprof["net"]], textposition="outside",
                hovertemplate="Age %{x}<br>Net internal migration: %{y:+,.0f}<extra></extra>"))
            figw.add_hline(y=0, line_color="#888888")
            figw.update_xaxes(title="Age band")
            figw.update_yaxes(title="Net internal migration (people)")
            show_chart(figw, "mig_profile", "ONS internal migration detailed estimates, 2024")

        # London ranking + map
        if borough_gj is not None:
            code_to_id, id_to_name = {}, {}
            for ft in borough_gj["features"]:
                p = ft["properties"]
                code = next((v for v in p.values() if isinstance(v, str) and v.startswith("E09")), None)
                nm = (p.get("name") or p.get("BoroughNa") or p.get("NAME") or
                      next((v for k, v in p.items() if isinstance(v, str) and "E09" not in v), ""))
                if code:
                    code_to_id[code] = ft["id"]; id_to_name[ft["id"]] = _norm_la(nm)
            dm2["gid"] = dm2["area_code"].map(code_to_id)
            dmap = dm2.dropna(subset=["gid"]).copy()
            dmap["name"] = dmap["gid"].map(id_to_name)
            if not dmap.empty:
                chart_title(f"{meas_m} of children aged {band_m} across London (2024)",
                            "Internal (within-UK) moves only · hover for the borough figure")
                if mcol == "net":
                    lim = float(np.nanmax(np.abs(dmap[mcol]))) or 1.0
                    scale = [[0, "#C0504D"], [0.5, "#F2F2F2"], [1, "#2E6E4E"]]
                    figmg = choropleth(borough_gj, dmap["gid"], dmap[mcol], dmap["name"],
                                       "Net migration", scale, fmt=":,.0f",
                                       zoom=9, center={"lat": 51.50, "lon": -0.12}, height=520)
                    figmg.data[0].update(zmin=-lim, zmax=lim)
                else:
                    figmg = choropleth(borough_gj, dmap["gid"], dmap[mcol], dmap["name"],
                                       meas_m, [[0, WCC["light_blue"]], [1, FOCAL]], fmt=":,.0f",
                                       zoom=9, center={"lat": 51.50, "lon": -0.12}, height=520)
                show_chart(figmg, "mig_map", "ONS internal migration detailed estimates, 2024")

                rank_m = dmap.sort_values(mcol)
                chart_title(f"London boroughs ranked — {meas_m.lower()}, ages {band_m} (2024)",
                            "Westminster highlighted")
                colr = np.where(rank_m["name"] == "Westminster", FOCAL, CONTEXT_BAR)
                figrm = go.Figure(go.Bar(
                    x=rank_m[mcol], y=rank_m["name"], orientation="h", marker_color=colr,
                    hovertemplate="<b>%{y}</b><br>" + meas_m + ": %{x:,.0f}<extra></extra>"))
                figrm.update_xaxes(title=f"{meas_m} (people)")
                figrm.update_yaxes(title="")
                figrm.update_layout(height=700)
                show_chart(figrm, "mig_rank", "ONS internal migration detailed estimates, 2024")

    st.divider()
    # ── 3 · School-age cohorts (single year of age) ──────────────────────────
    st.markdown("### 3 · Resident school-age cohorts")
    st.markdown(
        "Resident cohorts are built from **single years of age**, following the method used in the "
        "child-population analysis: the **primary** cohort is ages **4–10** (reception is age 4 "
        "turning 5; year 6 is age 10 turning 11) and the **secondary** cohort is ages **11–16** "
        "(every pupil turns 16 by the end of year 11). The transition bands (10–12 and 15–17) show "
        "where cohorts are lost between school phases.")
    if df_syoa.empty:
        st.info(
            "**Single-year-of-age population not loaded.** Add "
            "`population_single_year_of_age.csv` to the `data` folder (Nomis dataset NM_2002_1, "
            "ages 4–16 for Westminster and its neighbours); if it is absent the app calls the Nomis "
            "API directly. This section then shows each school cohort's size and its change over time, "
            "which is also the resident headcount behind the independent-school estimate.")
    else:
        q1, q2 = st.columns(2)
        coh = q1.selectbox("Cohort", list(SCHOOL_COHORTS), key="coh_sel")
        mode_c = q2.selectbox("Measure", ["% change vs first year", "Population"], key="coh_mode")
        lo, hi = SCHOOL_COHORTS[coh]
        dc = df_syoa[(df_syoa["age"] >= lo) & (df_syoa["age"] <= hi)]
        dc = dc.groupby(["year", "la"], as_index=False)["population"].sum()
        sel_c = st.multiselect("Boroughs", sorted(dc["la"].unique()),
                               default=sorted(dc["la"].unique()), key="coh_geos")
        dc = dc[dc["la"].isin(sel_c)]
        if dc.empty:
            st.info("No cohort data for that selection.")
        else:
            if mode_c.startswith("%"):
                ic = index_to_baseline(dc.rename(columns={"la": "geo", "population": "count"}),
                                       "count", int(dc["year"].min()), "pct_change")
                ic = ic.rename(columns={"geo": "la"}); ycol, ylab, fmt = "value", f"% change since {int(dc['year'].min())}", ":.1f"
            else:
                ic = dc.rename(columns={"population": "value"}); ycol, ylab, fmt = "value", "Resident children", ":,"
            chart_title(f"{coh} — {mode_c.lower()}",
                        "ONS mid-year estimates by single year of age · Westminster in strong colour")
            palcc = borough_palette(sorted(ic["la"].unique()))
            figcc = go.Figure()
            for g in sorted(ic["la"].unique()):
                sg = ic[ic["la"] == g].sort_values("year")
                figcc.add_trace(go.Scatter(
                    x=sg["year"], y=sg[ycol], mode="lines+markers", name=g,
                    line=dict(color=palcc[g], **line_style(g)),
                    marker=dict(size=8 if g == "Westminster" else 5),
                    hovertemplate="<b>" + g + "</b><br>%{x}: %{y" + fmt + "}<extra></extra>"))
            if mode_c.startswith("%"):
                figcc.add_hline(y=0, line_dash="dot", line_color="#BBBBBB")
            figcc.update_xaxes(title="Year"); figcc.update_yaxes(title=ylab)
            show_chart(figcc, "cohort_trend", "ONS mid-year estimates by single year of age (Nomis)")
            legend_hint()

    st.divider()
    # ── 4 · Cohort decline + projection ──────────────────────────────────────
    st.markdown("### 4 · The child population trend, and where it is heading")
    st.markdown(
        "The chart below tracks Westminster's resident child population and extends the recent "
        "trend forward. The projection is a **simple linear extrapolation** of the most recent "
        "years — it is a trend indicator, not an official ONS projection, and it assumes migration "
        "and birth patterns continue unchanged.")
    if df_mye_la.empty:
        st.info("Borough mid-year estimates not found — needed for the population trend.")
    else:
        p1, p2, p3 = st.columns(3)
        age_p = p1.selectbox("Age band", ["All 0–19"] + MYE_AGE_ORDER, key="proj_age")
        fit_p = p2.slider("Years of history used for the trend", 5, 20, 10, key="proj_fit")
        to_p = p3.slider("Project to", 2025, 2040, 2035, key="proj_to")
        sel_p = st.multiselect("Boroughs", list(NEIGHBOURS), default=["Westminster"], key="proj_geos")

        dpop = df_mye_la[(df_mye_la["gender"] == "Total") & (df_mye_la["area"].isin(sel_p))]
        if age_p != "All 0–19":
            dpop = dpop[dpop["age"] == age_p]
        dpop = dpop.groupby(["area", "year"], as_index=False)["population"].sum()
        if dpop.empty:
            st.info("No population data for that selection.")
        else:
            chart_title(f"Child population and projected trend to {to_p} — {age_p.lower()}",
                        "Solid = ONS mid-year estimates · dotted = linear projection of the recent trend")
            palp = borough_palette(sorted(dpop["area"].unique()))
            figpj = go.Figure()
            for g in sorted(dpop["area"].unique()):
                s = dpop[dpop["area"] == g].sort_values("year")
                figpj.add_trace(go.Scatter(
                    x=s["year"], y=s["population"], mode="lines", name=g,
                    line=dict(color=palp[g], **line_style(g)),
                    hovertemplate="<b>" + g + "</b><br>%{x}: %{y:,.0f}<extra></extra>"))
                hist = s.tail(fit_p)
                if len(hist) >= 3:
                    coef = np.polyfit(hist["year"], hist["population"], 1)
                    fx = np.arange(int(s["year"].max()), to_p + 1)
                    fy = np.polyval(coef, fx)
                    figpj.add_trace(go.Scatter(
                        x=fx, y=fy, mode="lines", name=f"{g} (projected)",
                        line=dict(color=palp[g], width=2, dash="dot"), showlegend=True,
                        hovertemplate="<b>" + g + " (projected)</b><br>%{x}: %{y:,.0f}<extra></extra>"))
            figpj.update_xaxes(title="Year")
            figpj.update_yaxes(title="Children", rangemode="tozero")
            show_chart(figpj, "pop_projection", "ONS mid-year estimates + linear trend projection")

            if "Westminster" in set(dpop["area"]):
                sw = dpop[dpop["area"] == "Westminster"].sort_values("year")
                hist = sw.tail(fit_p)
                if len(hist) >= 3:
                    coef = np.polyfit(hist["year"], hist["population"], 1)
                    proj = np.polyval(coef, to_p)
                    now = sw["population"].iloc[-1]
                    st.warning(
                        f"**On the trend of the last {fit_p} years**, Westminster's {age_p.lower()} "
                        f"population would fall from **{now:,.0f}** in {int(sw['year'].iloc[-1])} to "
                        f"about **{proj:,.0f}** by {to_p} — a change of **{(proj/now-1)*100:+.0f}%**. "
                        "Treat this as a direction of travel, not a forecast: it assumes the recent "
                        "rate of decline simply continues.")

    source_line("Births: ONS birth registrations via Nomis (see the User guide to birth statistics "
                "for definitions — births are assigned to the mother's usual residence). Internal "
                "migration: ONS detailed internal-migration estimates, 2024 (within-UK moves only; "
                "international migration is excluded). Population: ONS mid-year estimates. "
                "Projections are a simple linear extrapolation by this app, not an official projection.")

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.divider()
st.caption("WCC Children's Demographics · Westminster always shown in strong colour, comparators "
           "muted (Economist “grey-the-context” style) · every chart exports to PNG for slides · data "
           "sources linked in the sidebar.")
