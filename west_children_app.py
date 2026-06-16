# Westminster Children's Demographics Dashboard
# Data sources:
# 1. DWP Children in Low Income Families (AHC Relative) FYE 2025
# 2. DfE Key Stage 4 attainment — Explore Education Statistics
# 3. EGDI Ethnic Group Deprivation Index — Lloyd et al. 2023 / gedi.ac.uk
# 4. Census 2021 RM006, RM033 — ONS Nomis
# 5. Westminster LSOA + London borough boundaries — ONS / London Datastore

import os, io, json, warnings
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import topojson as tp
try:
    from pyproj import Transformer as _ProjTransformer
    _HAVE_PYPROJ = True
except ImportError:
    _HAVE_PYPROJ = False

warnings.filterwarnings("ignore")

# ── PATHS ─────────────────────────────────────────────────────────────────────
_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_DATA_SUBDIR = os.path.join(_SCRIPT_DIR, "data")
_SEARCH_DIRS = [_DATA_SUBDIR, _SCRIPT_DIR, "/mnt/user-data/uploads"]

_ALIASES = {
    "data-key-stage-4-performance__1_.csv": [
        "data-key-stage-4-performance__1_.csv",
        "data-key-stage-4-performance_1_.csv",
    ],
    "ONS_LSOA_2021 (1).json": [
        "ONS_LSOA_2021 (1).json",
        "ONS_LSOA_2021__1_.json",
        "ONS_LSOA_2021_(1).json",
    ],
}

def _dp(name):
    aliases = list(dict.fromkeys(_ALIASES.get(name, []) + [name]))
    for d in _SEARCH_DIRS:
        for a in aliases:
            p = os.path.join(d, a)
            if os.path.exists(p):
                return p
    return os.path.join("/mnt/user-data/uploads", name)

# ── WCC BRAND ─────────────────────────────────────────────────────────────────
WCC = {
    "blue":       "#0B2265",
    "yellow":     "#F5CB00",
    "cobalt":     "#0C35FA",
    "amaranth":   "#E34063",
    "green":      "#008466",
    "orange":     "#EA6F06",
    "white":      "#FFFFFF",
    "light_blue": "#E8EBF5",
    "grid":       "#E8E8E8",
}

BOROUGH_COLOURS = {
    "Westminster":            WCC["cobalt"],
    "Kensington & Chelsea":   WCC["blue"],
    "Camden":                 WCC["amaranth"],
    "Hammersmith & Fulham":   WCC["orange"],
    "Islington":              WCC["green"],
    "Wandsworth":             "#6B4C3B",
    "Kensington and Chelsea": WCC["blue"],
    "Hammersmith and Fulham": WCC["orange"],
}

NEIGHBOURS = {
    "Westminster":          "E09000033",
    "Kensington & Chelsea": "E09000020",
    "Camden":               "E09000007",
    "Hammersmith & Fulham": "E09000013",
    "Islington":            "E09000019",
    "Wandsworth":           "E09000032",
}

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Westminster Children's Demographics",
    page_icon="🏙️", layout="wide", initial_sidebar_state="expanded")

st.markdown(f"""<style>
  [data-testid="stSidebar"] {{
    background-color: {WCC["blue"]} !important;
    border-top: 4px solid {WCC["yellow"]};
  }}
  [data-testid="stSidebar"] * {{ color: #ffffff !important; }}
  [data-testid="stSidebar"] a {{ color: #A8C0FF !important; }}
  [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.25) !important; }}
  h1, h2, h3 {{ color: {WCC["blue"]} !important; font-family: Arial, sans-serif !important; }}
  .stTabs [data-baseweb="tab"] {{ font-size: .95rem; font-family: Arial, sans-serif; }}
  .stTabs [aria-selected="true"] {{
    color: {WCC["blue"]} !important;
    border-bottom: 3px solid {WCC["yellow"]} !important;
    font-weight: 700;
  }}
  [data-testid="stMetric"] {{
    background: {WCC["light_blue"]};
    border-radius: 6px;
    padding: 12px 14px;
    border-left: 4px solid {WCC["blue"]};
  }}
  .source-box {{
    background: {WCC["light_blue"]};
    border-radius: 5px;
    padding: 8px 12px;
    font-size: .82em;
    color: #333;
    margin-top: 6px;
    border-left: 3px solid {WCC["blue"]};
  }}
</style>""", unsafe_allow_html=True)

# ── HELPERS ───────────────────────────────────────────────────────────────────
def _pct(s):
    v = pd.to_numeric(s.astype(str).str.replace("%","",regex=False)
                       .str.replace(",","",regex=False).str.strip(), errors="coerce")
    if v.dropna().max() <= 1.0:
        v = v * 100
    return v

def _num(s):
    return pd.to_numeric(s.astype(str).str.replace(",","",regex=False).str.strip(), errors="coerce")

def apply_wcc_style(fig, source="", title_pad=10):
    """Apply WCC brand styling. Title should be in st.caption() above chart, not in plotly."""
    fig.update_layout(
        font_family="Arial", font_color=WCC["blue"],
        plot_bgcolor="white", paper_bgcolor="white",
        title={"text": ""},   # Explicitly empty — prevents 'undefined' in plotly 6.x
        margin=dict(l=50, r=30, t=title_pad, b=55),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, x=0,
            font=dict(family="Arial", size=11), title=dict(text="")
        ),
    )
    fig.update_xaxes(showgrid=False, linecolor="#cccccc", showline=True,
                     tickfont=dict(family="Arial", size=11))
    fig.update_yaxes(gridcolor=WCC["grid"], linecolor="white", zeroline=False,
                     tickfont=dict(family="Arial", size=11))
    if source:
        fig.add_annotation(
            text=f"<i>Source: {source}</i>",
            xref="paper", yref="paper", x=0, y=-0.16,
            showarrow=False, font=dict(size=9, color="#555", family="Arial"), align="left")
    return fig

def img_btn(fig, key):
    try:
        html_bytes = fig.to_html(include_plotlyjs="cdn", full_html=True).encode("utf-8")
        st.download_button(
            "⬇ Download chart (HTML)", data=html_bytes,
            file_name=f"{key}.html", mime="text/html", key=f"dl_{key}")
    except Exception as e:
        st.caption(f"_Download unavailable: {e}_")

def highlight_westminster(fig, df, x_col, y_col, name_col="Borough", name="Westminster",
                          marker_size=14, color=None, label=None):
    """Add a highlighted marker + annotation for Westminster on any scatter/line chart."""
    wcc_rows = df[df[name_col].str.contains(name, na=False) | df.get(name_col, pd.Series()).eq(name)]
    if hasattr(df, name_col) and name_col in df.columns:
        wcc_rows = df[df[name_col] == name]
    if wcc_rows.empty:
        return fig
    x_val = wcc_rows[x_col].iloc[0]
    y_val = wcc_rows[y_col].iloc[0]
    lbl   = label or f"Westminster: {y_val:.1f}"
    fig.add_trace(go.Scatter(
        x=[x_val], y=[y_val], mode="markers+text",
        marker=dict(size=marker_size, color=color or WCC["yellow"],
                    line=dict(width=2, color=WCC["blue"])),
        text=[lbl], textposition="top center",
        textfont=dict(color=WCC["blue"], size=11, family="Arial"),
        name="Westminster", showlegend=False
    ))
    return fig

def _card(col, title, value, sub, help_txt):
    col.metric(title, value, help=help_txt)
    col.markdown(f"<div style='font-size:.78em;color:#666;margin-top:-14px'>{sub}</div>",
                 unsafe_allow_html=True)

# ── MAP HELPERS — inject 'id' for go.Choroplethmap ───────────────────────────
def _inject_id(geojson, id_property):
    """Return a copy of GeoJSON FeatureCollection with 'id' set from a property."""
    out = {"type": "FeatureCollection", "features": []}
    for f in geojson["features"]:
        fc = {k: v for k, v in f.items()}
        fc["id"] = f["properties"][id_property]
        out["features"].append(fc)
    return out

def _reproject_osgb_to_wgs84(geojson):
    """Reproject GeoJSON coordinates from OSGB EPSG:27700 to WGS84 EPSG:4326.
    Required when the GeoJSON uses British National Grid eastings/northings.
    """
    if not _HAVE_PYPROJ:
        st.warning("pyproj not installed — LSOA map coordinates cannot be reprojected. Add 'pyproj' to requirements.txt.")
        return geojson
    import copy
    tr = _ProjTransformer.from_crs("EPSG:27700", "EPSG:4326", always_xy=True)
    out = copy.deepcopy(geojson)
    for feat in out["features"]:
        geom = feat["geometry"]
        if geom["type"] == "Polygon":
            geom["coordinates"] = [
                [list(tr.transform(x, y)) for x, y in ring]
                for ring in geom["coordinates"]
            ]
        elif geom["type"] == "MultiPolygon":
            geom["coordinates"] = [
                [[list(tr.transform(x, y)) for x, y in ring] for ring in poly]
                for poly in geom["coordinates"]
            ]
    return out

def lsoa_choropleth(geojson, codes, z_vals, names, label, colorscale, height=540):
    """Build a go.Choroplethmap for Westminster LSOAs."""
    fig = go.Figure(go.Choroplethmap(
        geojson=geojson,
        locations=codes,
        z=z_vals,
        text=names,
        hovertemplate="<b>%{text}</b><br>" + label + ": %{z:.1f}<extra></extra>",
        colorscale=colorscale,
        marker_opacity=0.75,
        marker_line_width=0.3,
        colorbar=dict(title=dict(text=label, font=dict(size=11)), thickness=14, len=0.6),
    ))
    fig.update_layout(
        map_style="carto-positron",
        map_zoom=12, map_center={"lat": 51.512, "lon": -0.155},
        margin=dict(l=0, r=0, t=0, b=0), height=height,
        paper_bgcolor="white",
    )
    return fig

def borough_choropleth(geojson, codes, z_vals, names, label, colorscale,
                       zoom=10.5, center=None, height=460):
    """Build a go.Choroplethmap for London boroughs."""
    if center is None:
        center = {"lat": 51.505, "lon": -0.17}
    fig = go.Figure(go.Choroplethmap(
        geojson=geojson,
        locations=codes,
        z=z_vals,
        text=names,
        hovertemplate="<b>%{text}</b><br>" + label + ": %{z:.1f}%<extra></extra>",
        colorscale=colorscale,
        marker_opacity=0.75, marker_line_width=1,
        zmin=min(z_vals) * 0.9, zmax=max(z_vals) * 1.05,
        colorbar=dict(title=dict(text=label, font=dict(size=11)), thickness=14, len=0.6),
    ))
    fig.update_layout(
        map_style="carto-positron",
        map_zoom=zoom, map_center=center,
        margin=dict(l=0, r=0, t=0, b=0), height=height,
        paper_bgcolor="white",
    )
    return fig

# ── DATA LOADERS ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_low_income_la():
    df = pd.read_csv(_dp("2_AHC_Relative_LA.csv"), header=8,
                     names=["LA","Area_Code","N_2024","N_2025","Pct_2024","Pct_2025"],
                     usecols=[0,1,2,3,4,5])
    df = df.dropna(subset=["Area_Code"])
    df["Pct_2024"] = _pct(df["Pct_2024"])
    df["Pct_2025"] = _pct(df["Pct_2025"])
    df["N_2024"]   = _num(df["N_2024"])
    df["N_2025"]   = _num(df["N_2025"])
    return df

@st.cache_data(show_spinner=False)
def load_low_income_ward():
    df = pd.read_csv(_dp("4_AHC_Relative_Ward.csv"), header=9,
                     names=["LA","LA_Code","Ward","Ward_Code","N_2024","N_2025","Pct_2024","Pct_2025"],
                     usecols=[0,1,2,3,4,5,6,7])
    df = df.dropna(subset=["Ward_Code"])
    df["Pct_2024"] = _pct(df["Pct_2024"])
    df["Pct_2025"] = _pct(df["Pct_2025"])
    df["N_2024"]   = _num(df["N_2024"])
    df["N_2025"]   = _num(df["N_2025"])
    df["LA_filled"] = df["LA"].ffill()
    return df

@st.cache_data(show_spinner=False)
def load_ks4_ethnic():
    """KS4 by ethnic group 2024/25.
    File: data-key-stage-4-performance.csv
    col0=ethnic_group, col1=subgroup, col2=region, col3=LA, col4+=data
    row2=year(ffill), row3=metric(ffill), row4=Total/FSM, data from row5.
    """
    path = _dp("data-key-stage-4-performance.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["ethnic_group","subgroup","la","att8_2425"])
    df = pd.read_csv(path, header=None, low_memory=False)

    r2 = df.iloc[2, 4:].astype(str).ffill()
    r3 = df.iloc[3, 4:].astype(str).ffill()
    r4 = df.iloc[4, 4:].astype(str).ffill()

    cols = ["ethnic_group","subgroup","region","la"]
    seen = {}
    for j in range(len(r2)):
        yr  = r2.iloc[j].strip().replace("/","_")
        m   = r3.iloc[j].strip().replace(" ","_")[:22]
        s   = r4.iloc[j].strip().replace(" ","_")[:18]
        base = f"{yr}_{m}_{s}"
        seen[base] = seen.get(base, 0) + 1
        cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")

    data = df.iloc[5:].copy()
    data.columns = cols[:len(data.columns)]
    for c in ["ethnic_group","subgroup","region"]:
        data[c] = data[c].ffill()

    inner = ["Camden","Hackney","Hammersmith and Fulham","Haringey","Islington",
             "Kensington and Chelsea","Lambeth","Lewisham","Newham","Southwark",
             "Tower Hamlets","Wandsworth","Westminster"]

    mask = data["la"].isin(inner) & data["subgroup"].astype(str).str.startswith("All")
    sub  = data[mask].copy()

    att_col = [c for c in sub.columns if "2024_25" in c and "Attainment_8" in c and "Total" in c]
    pct_col = [c for c in sub.columns if "2024_25" in c and "achieving_gr" in c and "Total" in c]

    sub["att8_2425"] = pd.to_numeric(
        sub[att_col[0]].astype(str).str.replace("no data","",regex=False),
        errors="coerce") if att_col else np.nan
    sub["pct_5above"] = pd.to_numeric(
        sub[pct_col[0]].astype(str).str.replace("no data","",regex=False).str.replace("%","",regex=False),
        errors="coerce") if pct_col else np.nan

    sub = sub.copy()
    sub["la"] = (sub["la"].astype(str)
                 .str.replace("and Fulham","& Fulham",regex=False)
                 .str.replace("and Chelsea","& Chelsea",regex=False))
    sub["ethnic_group"] = (sub["ethnic_group"].astype(str)
                           .str.replace("Asian / Asian British","Asian",regex=False)
                           .str.replace("Black / African / Caribbean / Black British","Black",regex=False)
                           .str.replace("Mixed / multiple ethnic groups","Mixed",regex=False)
                           .str.replace("Other ethnic group","Other",regex=False))
    return sub

@st.cache_data(show_spinner=False)
def load_ks4_time():
    """KS4 time series all-pupils 2018/19–2024/25.
    File: data-key-stage-4-performance__1_.csv
    col0=ethnic_group, col1=subgroup, col2=sex, col3=region, col4=LA
    row2=metric(ffill), row3=year, data from row5.
    All-pupils = NaN in cols 0,1,2. Att8 Total at cols 12-18.
    """
    path = _dp("data-key-stage-4-performance__1_.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["la","year","att8"])
    df = pd.read_csv(path, header=None, low_memory=False)

    YEARS     = ["2018/19","2019/20","2020/21","2021/22","2022/23","2023/24","2024/25"]
    ATT8_COLS = [12,13,14,15,16,17,18]

    df_c = df.copy()
    for ci in [0,1,2,3]:
        df_c.iloc[5:, ci] = df_c.iloc[5:, ci].ffill()

    inner = {"Camden","Hackney","Hammersmith and Fulham","Haringey","Islington",
             "Kensington and Chelsea","Lambeth","Lewisham","Newham","Southwark",
             "Tower Hamlets","Wandsworth","Westminster"}

    sub = df_c.iloc[5:].copy()
    mask = (sub.iloc[:,4].isin(inner) &
            sub.iloc[:,0].isna() & sub.iloc[:,1].isna() & sub.iloc[:,2].isna())
    all_pupils = sub[mask]

    records = []
    for _, row in all_pupils.iterrows():
        la = str(row.iloc[4]).strip()
        for yr, cidx in zip(YEARS, ATT8_COLS):
            val = pd.to_numeric(str(row.iloc[cidx]).replace(",","").strip(), errors="coerce")
            records.append({"la": la, "year": yr, "att8": val})

    ts = pd.DataFrame(records) if records else pd.DataFrame(columns=["la","year","att8"])
    if not ts.empty:
        ts["la"] = (ts["la"].str.replace("and Fulham","& Fulham",regex=False)
                             .str.replace("and Chelsea","& Chelsea",regex=False))
    return ts

@st.cache_data(show_spinner=False)
def load_rm006():
    df = pd.read_excel(
        _dp("RM006_age_of_youngest_dependent_child_by_household_type.xlsx"),
        header=7, skiprows=[8])
    df.columns = ["LSOA","No_dep_children","Age_0_4","Age_5_9","Age_10_15","Age_16_18"]
    df = df.dropna(subset=["LSOA"])
    df["LSOA_CODE"] = df["LSOA"].astype(str).str.extract(r"(E\d+)")
    df["LSOA_NAME"] = df["LSOA"].astype(str).str.replace(r"E\d+ : ","",regex=True).str.strip()
    for c in ["Age_0_4","Age_5_9","Age_10_15","Age_16_18","No_dep_children"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    df["Total_dep_children"] = df[["Age_0_4","Age_5_9","Age_10_15","Age_16_18"]].sum(axis=1)
    df["Pct_under5"] = np.where(df["Total_dep_children"]>0,
                                df["Age_0_4"]/df["Total_dep_children"]*100, 0)
    return df[df["LSOA_CODE"].notna()].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_rm033():
    df = pd.read_excel(_dp("RM033_ethic_group_dependent_child_by_sex.xlsx"), header=8)
    df.rename(columns={df.columns[0]:"LSOA"}, inplace=True)
    df = df.dropna(subset=["LSOA"])
    df["LSOA_CODE"] = df["LSOA"].astype(str).str.extract(r"(E\d+)")
    df["LSOA_NAME"] = df["LSOA"].astype(str).str.replace(r"E\d+ : ","",regex=True).str.strip()
    eth = [c for c in df.columns if c not in ["LSOA","LSOA_CODE","LSOA_NAME"]]
    for c in eth:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Asian"] = df[[c for c in eth if "Asian" in c]].sum(axis=1)
    df["Black"]  = df[[c for c in eth if "Black" in c]].sum(axis=1)
    df["Mixed"]  = df[[c for c in eth if "Mixed" in c]].sum(axis=1)
    df["White"]  = df[[c for c in eth if "White" in c]].sum(axis=1)
    df["Arab"]   = df[[c for c in eth if "Arab"  in c]].sum(axis=1)
    df["Other"]  = df[[c for c in eth if ("Other" in c and "Black" not in c
                       and "Asian" not in c and "White" not in c and "Mixed" not in c)]].sum(axis=1)
    df["Total"]  = df[["Asian","Black","Mixed","White","Arab","Other"]].sum(axis=1)
    return df[df["LSOA_CODE"].notna()].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_egdi_lsoa():
    """LSOA-level EGDI data. Returns empty DF if file not found."""
    path = _dp("EGDI.xlsx")
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name="Data")
    edi_cols = [c for c in df.columns if c.startswith("EDI.")]
    keep = ["LSOA21CD","LSOA21NM","Range","Mostdeprivedgroup","Leastdeprivedgroup",
            "TopEGDIDEC","BottomEGDIDEC"] + edi_cols
    return df[[c for c in keep if c in df.columns]].copy()

@st.cache_data(show_spinner=False)
def load_egdi():
    df = pd.read_excel(_dp("EGDI-Local-Authority-profiles.xlsx"), sheet_name="Profiles")
    df.columns = ["idx","LA_Code","LA_Name",
                  "D1","D2","D3","D4","D5","D6","D7","D8","D9","D10",
                  "Total_LSOAs","Pct_D1","Pct_D2","Pct_D3","Pct_D4","Pct_D5",
                  "Pct_D6","Pct_D7","Pct_D8","Pct_D9","Pct_D10",
                  "_a","_b","_c","Category","_d","Flat","More_ethnic_ineq",
                  "Less_ethnic_ineq","N_shape","Pct_bottom20","Pct_top20"]
    return df.iloc[1:].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_wcc_geojson():
    """Westminster 2021 LSOA boundaries — reprojected to WGS84, with 'id' injected."""
    with open(_dp("ONS_LSOA_2021 (1).json")) as f:
        gj = json.load(f)
    # Check if coordinates are in OSGB (British National Grid) — values >> 1000
    sample_coord = gj["features"][0]["geometry"]["coordinates"][0][0]
    if abs(sample_coord[0]) > 1000:  # OSGB eastings are ~500000
        gj = _reproject_osgb_to_wgs84(gj)
    return _inject_id(gj, "LSOA21CD")

@st.cache_data(show_spinner=False)
def load_borough_geojson():
    """London borough boundaries from TopoJSON, with injected 'id' field."""
    with open(_dp("Borough_London_LL84.json")) as f:
        topo = json.load(f)
    gj = json.loads(tp.Topology(topo, object_name="Borough_London_LL84").to_geojson())
    return _inject_id(gj, "BoroughCod")

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = _dp("city_of_westminster.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=160)
    else:
        st.image("https://raw.githubusercontent.com/SianReilly/childrens_demographics_wcc/main/data/city_of_westminster.png", width=160)
    st.markdown("## Westminster Children's Dashboard")
    st.markdown("**CIPFA Statistical Neighbours**")
    for b, code in NEIGHBOURS.items():
        col = BOROUGH_COLOURS.get(b, WCC["cobalt"])
        st.markdown(f"<span style='color:{col}'>■</span> {b}", unsafe_allow_html=True)
    st.divider()
    st.caption("**Data sources**")
    st.markdown("[DWP Children in Low Income Families](https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025)")
    st.markdown("[DfE Key Stage 4 (Explore Ed Stats)](https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df)")
    st.markdown("[EGDI — gedi.ac.uk](https://gedi.ac.uk/egdi/)")
    st.markdown("[Census 2021 — ONS Nomis](https://www.nomisweb.co.uk/)")
    st.divider()
    st.caption("Census 2021 · KS4 2024/25 · Child poverty FYE 2025")

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
with st.spinner("Loading datasets…"):
    df_li_la       = load_low_income_la()
    df_li_ward     = load_low_income_ward()
    df_ks4_eth     = load_ks4_ethnic()
    df_ks4_ts      = load_ks4_time()
    df_rm006       = load_rm006()
    df_rm033       = load_rm033()
    df_egdi        = load_egdi()
    df_egdi_lsoa   = load_egdi_lsoa()
    wcc_geojson    = load_wcc_geojson()
    borough_geojson = load_borough_geojson()

# ── HEADER & METRICS ──────────────────────────────────────────────────────────
st.title("🏙️ Westminster Children's Demographics")
st.markdown("Child poverty, demographics, attainment and ethnic diversity — benchmarked against CIPFA statistical neighbours.")

wcc_li_rows = df_li_la[df_li_la["LA"].str.contains("Westminster", na=False)]
wcc_li = wcc_li_rows.iloc[0] if len(wcc_li_rows) else pd.Series(
    {"N_2024":0,"N_2025":0,"Pct_2024":0,"Pct_2025":0})
london_avg   = df_li_la[df_li_la["Area_Code"].astype(str).str.startswith("E09", na=False)]["Pct_2025"].mean()
total_dep    = df_rm006["Total_dep_children"].sum()
pct_u5       = round(df_rm006["Age_0_4"].sum()/total_dep*100, 1) if total_dep else 0
rm_tot       = df_rm033["Total"].sum()
pct_nonwhite = round((1 - df_rm033["White"].sum()/rm_tot)*100, 1) if rm_tot else 0

c1, c2, c3, c4, c5 = st.columns(5)
_card(c1, "Children in low income (FYE 2025)", f"{int(wcc_li['N_2025']):,}",
      f"{wcc_li['Pct_2025']:.1f}% of all Westminster children",
      "AHC relative poverty, children aged 0–15. FYE = financial year ending. Source: DWP/HMRC.")

c2.metric("Change vs FYE 2024",
          f"{int(wcc_li['N_2025']-wcc_li['N_2024']):+,} children",
          delta=f"{wcc_li['Pct_2025']-wcc_li['Pct_2024']:+.1f}pp year-on-year",
          delta_color="inverse",
          help="Year-on-year change FYE 2024→2025. Green ↓ = fewer children in poverty (improvement). Red ↑ = deterioration.")

diff_lon = wcc_li['Pct_2025'] - london_avg
c3.metric("vs London borough average",
          f"{wcc_li['Pct_2025']:.1f}% (Westminster)",
          delta=f"{diff_lon:+.1f}pp vs avg of 33 boroughs ({london_avg:.1f}%)",
          delta_color="inverse",
          help=f"Westminster vs unweighted average of all 33 London boroughs ({london_avg:.1f}%). Green ↓ = below London average.")

_card(c4, "Dependent children (Census 2021)", f"{total_dep:,}",
      f"{pct_u5}% aged 0–4 · Census 2021 snapshot",
      "Total dependent children across Westminster's 123 LSOAs (2021 boundaries). Census 2021, ONS Nomis RM006.")

_card(c5, "Non-white children (Census 2021)", f"{pct_nonwhite}%",
      "of dependent children in Westminster",
      "% of dependent children not classified as White. Census 2021, ONS Nomis RM033.")

st.divider()

# ── TABS ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📍 Child Poverty",
    "🗺️ Westminster LSOA Maps",
    "📚 KS4 Attainment",
    "👥 Ethnicity & Demographics",
    "⚖️ Ethnic Deprivation (EGDI)",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — CHILD POVERTY
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Children in low income families — Westminster vs CIPFA neighbours (FYE 2025)")
    st.caption("After-housing-costs (AHC) relative poverty, children aged 0–15. Source: DWP/HMRC.")

    nb_codes = list(NEIGHBOURS.values())
    df_nb = df_li_la[df_li_la["Area_Code"].isin(nb_codes)].copy()
    df_nb["Borough"] = (df_nb["LA"]
                        .str.replace("and Fulham","& Fulham")
                        .str.replace("and Chelsea","& Chelsea"))
    df_nb = df_nb.sort_values("Pct_2025", ascending=True)

    col_a, col_b = st.columns([3,2])
    with col_a:
        fig1 = px.bar(df_nb, x="Pct_2025", y="Borough", orientation="h",
                      color="Borough",
                      color_discrete_map={r["Borough"]: BOROUGH_COLOURS.get(r["Borough"], WCC["blue"])
                                          for _, r in df_nb.iterrows()},
                      text="Pct_2025")
        fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        # Highlight Westminster bar with yellow outline
        wcc_idx = df_nb["Borough"].tolist().index("Westminster") if "Westminster" in df_nb["Borough"].tolist() else -1
        if wcc_idx >= 0:
            fig1.add_shape(type="rect",
                y0=wcc_idx-0.4, y1=wcc_idx+0.4,
                x0=0, x1=df_nb["Pct_2025"].iloc[wcc_idx],
                line=dict(color=WCC["yellow"], width=3), fillcolor="rgba(0,0,0,0)",
                yref="y", xref="x")
        fig1.update_xaxes(range=[0, df_nb["Pct_2025"].max()*1.22], title="% children in low income (AHC)")
        fig1.update_yaxes(title="")
        apply_wcc_style(fig1, "DWP Children in Low Income Families FYE 2025")
        st.caption("Islington highest; Kensington & Chelsea lowest — Westminster highlighted")
        st.plotly_chart(fig1, use_container_width=True)
        img_btn(fig1, "child_poverty_bar")

    with col_b:
        df_ts2 = df_nb[["Borough","Pct_2024","Pct_2025"]].melt("Borough", var_name="Year", value_name="Pct")
        df_ts2["Year"] = df_ts2["Year"].map({"Pct_2024":"FYE 2024","Pct_2025":"FYE 2025"})
        fig2 = px.line(df_ts2, x="Year", y="Pct", color="Borough", markers=True,
                       color_discrete_map={b: BOROUGH_COLOURS.get(b, WCC["blue"])
                                          for b in df_ts2["Borough"].unique()})
        fig2.update_traces(line_width=2.5)
        # Make Westminster line thicker and add yellow markers
        for trace in fig2.data:
            if hasattr(trace, 'name') and trace.name == "Westminster":
                trace.line.width = 4
                trace.marker.size = 12
                trace.marker.color = WCC["yellow"]
                trace.marker.line = dict(width=2, color=WCC["blue"])
        fig2.update_yaxes(title="% children in low income", rangemode="tozero")
        fig2.update_xaxes(title="")
        apply_wcc_style(fig2, "DWP Children in Low Income Families FYE 2025")
        st.caption("All neighbours saw child poverty fall 2024→2025 — Westminster highlighted")
        st.plotly_chart(fig2, use_container_width=True)
        img_btn(fig2, "child_poverty_trend")

    # ── CIPFA choropleth map
    st.divider()
    st.subheader("Geographic context — CIPFA statistical neighbours")
    st.caption("% children in low income (AHC relative, FYE 2025). Hover for value.")

    nb_gj_filtered = {
        "type": "FeatureCollection",
        "features": [f for f in borough_geojson["features"] if f["id"] in nb_codes]
    }
    fig_nb_map = borough_choropleth(
        nb_gj_filtered,
        codes=df_nb["Area_Code"].tolist(),
        z_vals=df_nb["Pct_2025"].tolist(),
        names=df_nb["Borough"].tolist(),
        label="% in low income",
        colorscale=[[0,"#E8EBF5"],[0.5,WCC["cobalt"]],[1.0,WCC["blue"]]],
    )
    st.plotly_chart(fig_nb_map, use_container_width=True)
    img_btn(fig_nb_map, "cipfa_map")

    st.divider()
    st.subheader("Westminster ward-level child poverty (FYE 2025)")
    wcc_wards = df_li_ward[df_li_ward["LA_filled"].astype(str).str.contains("Westminster",na=False)].copy()
    wcc_wards = wcc_wards.sort_values("Pct_2025", ascending=True).dropna(subset=["Ward","Pct_2025"])
    fig3 = px.bar(wcc_wards, x="Pct_2025", y="Ward", orientation="h",
                  color_discrete_sequence=[WCC["blue"]], text="Pct_2025")
    fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig3.update_xaxes(range=[0, wcc_wards["Pct_2025"].max()*1.22],
                      title="% children in low income (AHC)")
    fig3.update_yaxes(title="")
    apply_wcc_style(fig3, "DWP Children in Low Income Families FYE 2025 — ward level")
    st.caption("Church Street & Westbourne have highest child poverty in Westminster")
    st.plotly_chart(fig3, use_container_width=True)
    img_btn(fig3, "ward_poverty")

    st.markdown('<div class="source-box">Child poverty figures use the AHC (after-housing-costs) relative measure, children aged 0–15. FYE = financial year ending. AHC is generally considered the more meaningful measure as it accounts for housing cost variation across London. Source: DWP/HMRC Children in Low Income Families, published 2025.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — LSOA MAPS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Westminster LSOA demographic maps (Census 2021)")
    st.caption("123 LSOAs using 2021 ONS boundaries. Hover for LSOA name and value.")

    map_metric = st.selectbox("Map metric", [
        "Total dependent children","% aged 0–4","% aged 10–15",
        "% White children","% Asian children","% Black children","% Mixed children"])

    df_map = pd.merge(
        df_rm006[["LSOA_CODE","LSOA_NAME","Total_dep_children","Age_0_4","Age_10_15"]],
        df_rm033[["LSOA_CODE","White","Asian","Black","Mixed","Total"]],
        on="LSOA_CODE", how="inner")
    agg = df_map.groupby("LSOA_CODE").agg(
        LSOA_NAME=("LSOA_NAME","first"),
        Total=("Total_dep_children","sum"),
        Age_0_4=("Age_0_4","sum"), Age_10_15=("Age_10_15","sum"),
        White=("White","sum"), Asian=("Asian","sum"),
        Black=("Black","sum"), Mixed=("Mixed","sum"), Eth_Total=("Total","sum")
    ).reset_index()
    sd = lambda n, d: np.where(d > 0, n/d*100, 0)
    agg["pct_u5"]    = sd(agg["Age_0_4"],   agg["Total"])
    agg["pct_10_15"] = sd(agg["Age_10_15"], agg["Total"])
    agg["pct_white"] = sd(agg["White"],     agg["Eth_Total"])
    agg["pct_asian"] = sd(agg["Asian"],     agg["Eth_Total"])
    agg["pct_black"] = sd(agg["Black"],     agg["Eth_Total"])
    agg["pct_mixed"] = sd(agg["Mixed"],     agg["Eth_Total"])

    METRIC_MAP = {
        "Total dependent children": ("Total",     "Total dependent children", "Blues"),
        "% aged 0–4":               ("pct_u5",    "% aged 0–4",              "YlOrBr"),
        "% aged 10–15":             ("pct_10_15", "% aged 10–15",            "Purples"),
        "% White children":         ("pct_white", "% White children",         "Greys"),
        "% Asian children":         ("pct_asian", "% Asian children",         "Oranges"),
        "% Black children":         ("pct_black", "% Black children",         "Reds"),
        "% Mixed children":         ("pct_mixed", "% Mixed children",         "Greens"),
    }
    col_name, label, cscale = METRIC_MAP[map_metric]

    # Match agg to geojson codes
    geojson_codes = [f["id"] for f in wcc_geojson["features"]]
    agg_indexed = agg.set_index("LSOA_CODE")
    matched_codes = [c for c in geojson_codes if c in agg_indexed.index]
    matched_z     = [float(agg_indexed.loc[c, col_name]) for c in matched_codes]
    matched_names = [str(agg_indexed.loc[c, "LSOA_NAME"]) for c in matched_codes]

    st.markdown(f"**{label} by Westminster LSOA (Census 2021)**")
    fig_map = lsoa_choropleth(wcc_geojson, matched_codes, matched_z, matched_names,
                              label, cscale, height=540)
    st.plotly_chart(fig_map, use_container_width=True)
    img_btn(fig_map, f"lsoa_{col_name}")

    st.markdown('<div class="source-box">Census 2021, ONS Nomis. RM006: age of youngest dependent child by household type. RM033: ethnic group of dependent child by sex. 2021 LSOA boundaries (123 LSOAs).</div>', unsafe_allow_html=True)

    st.divider()
    top15 = agg.nlargest(15,"Total").sort_values("Total")
    fig_lsoa = px.bar(top15, x="Total", y="LSOA_NAME", orientation="h",
                      color_discrete_sequence=[WCC["blue"]], text="Total")
    fig_lsoa.update_traces(textposition="outside")
    fig_lsoa.update_xaxes(title="Total dependent children (Census 2021)")
    fig_lsoa.update_yaxes(title="")
    apply_wcc_style(fig_lsoa, "Census 2021, ONS Nomis RM006")
    st.caption("Top 15 Westminster LSOAs by number of dependent children")
    st.plotly_chart(fig_lsoa, use_container_width=True)
    img_btn(fig_lsoa, "lsoa_top15")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — KS4 ATTAINMENT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Key Stage 4 attainment — Westminster vs CIPFA neighbours")
    st.caption("Average Attainment 8 score, 2024/25, state-funded schools. Source: DfE Explore Education Statistics.")

    nb_la  = ["Camden","Hammersmith & Fulham","Islington","Kensington & Chelsea","Wandsworth","Westminster"]
    ks4_nb = df_ks4_eth[df_ks4_eth["la"].isin(nb_la) & df_ks4_eth["att8_2425"].notna()].copy()

    if ks4_nb.empty:
        st.info("KS4 ethnic data not available — ensure `data-key-stage-4-performance.csv` is in the `data/` folder.")
    else:
        col_ka, col_kb = st.columns(2)
        with col_ka:
            wcc_eth = (ks4_nb[ks4_nb["la"]=="Westminster"]
                       .dropna(subset=["att8_2425"])
                       .drop_duplicates("ethnic_group")
                       .sort_values("att8_2425"))
            fig_ae = px.bar(wcc_eth, x="att8_2425", y="ethnic_group", orientation="h",
                            color_discrete_sequence=[WCC["blue"]], text="att8_2425")
            fig_ae.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            # Highlight the 'All ... pupils' bar (top summary row) in yellow
            all_idx_list = [i for i,g in enumerate(wcc_eth["ethnic_group"].tolist())
                            if "all" in str(g).lower()]
            if all_idx_list:
                ai = all_idx_list[0]
                wcc_eth_list = wcc_eth["ethnic_group"].tolist()
                fig_ae.add_shape(type="rect",
                    y0=ai-0.4, y1=ai+0.4,
                    x0=0, x1=float(wcc_eth["att8_2425"].iloc[ai]),
                    line=dict(color=WCC["yellow"], width=3), fillcolor="rgba(0,0,0,0)",
                    yref="y", xref="x")
            fig_ae.update_xaxes(title="Average Attainment 8 score", range=[0, 80])
            fig_ae.update_yaxes(title="")
            apply_wcc_style(fig_ae, "DfE KS4 2024/25, state-funded schools")
            st.caption("Westminster: Attainment 8 by ethnic group (2024/25)")
            st.plotly_chart(fig_ae, use_container_width=True)
            img_btn(fig_ae, "wcc_att8_ethnic")

        with col_kb:
            # Attainment 8 comparison across CIPFA neighbours — all pupils only
            all_nb = (ks4_nb[ks4_nb["ethnic_group"].str.lower().str.contains("all", na=False) |
                             ks4_nb["subgroup"].str.lower().str.contains("all pupils", na=False)]
                      .drop_duplicates("la")
                      .sort_values("att8_2425"))
            if not all_nb.empty:
                fig_nb_att = px.bar(all_nb, x="att8_2425", y="la", orientation="h",
                                    color="la",
                                    color_discrete_map={la: BOROUGH_COLOURS.get(la, WCC["blue"]) for la in all_nb["la"]},
                                    text="att8_2425")
                fig_nb_att.update_traces(texttemplate="%{text:.1f}", textposition="outside", showlegend=False)
                fig_nb_att.update_xaxes(title="Avg Attainment 8 score", range=[0,70])
                fig_nb_att.update_yaxes(title="")
                apply_wcc_style(fig_nb_att, "DfE KS4 2024/25, state-funded schools")
                st.caption("All-pupils Attainment 8 — CIPFA neighbours (2024/25)")
                st.plotly_chart(fig_nb_att, use_container_width=True)
                img_btn(fig_nb_att, "cipfa_att8")



    # ── Time series
    st.divider()
    st.subheader("Attainment 8 trend 2018/19–2024/25")
    if df_ks4_ts.empty or "la" not in df_ks4_ts.columns:
        st.info("Time series data not available — ensure `data-key-stage-4-performance__1_.csv` is in `data/`.")
    else:
        nb_v1 = ["Camden","Hammersmith & Fulham","Islington","Kensington & Chelsea","Wandsworth","Westminster"]
        ts_d = df_ks4_ts[df_ks4_ts["la"].isin(nb_v1) & df_ks4_ts["att8"].notna()].copy()
        if not ts_d.empty:
            fig_ts = px.line(ts_d, x="year", y="att8", color="la", markers=True,
                             color_discrete_map={la: BOROUGH_COLOURS.get(la, WCC["blue"])
                                               for la in ts_d["la"].unique()})
            for trace in fig_ts.data:
                if hasattr(trace, 'name') and trace.name == "Westminster":
                    trace.line.width = 4
                    trace.marker.size = 10
                    trace.marker.color = WCC["yellow"]
                    trace.marker.line = dict(width=2, color=WCC["blue"])
            fig_ts.update_traces(line_width=2.5)
            fig_ts.update_xaxes(title="Academic year")
            fig_ts.update_yaxes(title="Avg Attainment 8 score", rangemode="tozero")
            apply_wcc_style(fig_ts, "DfE KS4 Performance, state-funded schools, Explore Education Statistics")
            st.caption("Attainment 8 dipped 2020/21–2021/22 (COVID disruption) then recovered — Westminster highlighted")
            st.plotly_chart(fig_ts, use_container_width=True)
            img_btn(fig_ts, "att8_trend")

    st.markdown('<div class="source-box">DfE Key Stage 4 attainment by ethnicity, state-funded schools, Inner London. 2024/25. Note: 2020/21 and 2021/22 figures should be interpreted with caution due to COVID-19 assessment disruption. Small ethnic group counts lead to suppressed cells (shown as no data).</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ETHNICITY & DEMOGRAPHICS
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Children's ethnicity & age structure — Westminster LSOAs (Census 2021)")
    col_e1, col_e2 = st.columns(2)

    with col_e1:
        eth_totals = {"Asian":   df_rm033["Asian"].sum(),
                      "Black":   df_rm033["Black"].sum(),
                      "Mixed":   df_rm033["Mixed"].sum(),
                      "White":   df_rm033["White"].sum(),
                      "Arab":    df_rm033["Arab"].sum(),
                      "Other":   df_rm033["Other"].sum()}
        eth_df = pd.DataFrame(list(eth_totals.items()), columns=["Group","Count"])
        eth_df = eth_df.sort_values("Count", ascending=True)
        eth_df["Pct"] = eth_df["Count"] / eth_df["Count"].sum() * 100
        fig_eb = px.bar(eth_df, x="Count", y="Group", orientation="h",
                        color="Group",
                        color_discrete_sequence=[WCC["blue"],WCC["cobalt"],WCC["amaranth"],
                                                 WCC["yellow"],WCC["green"],WCC["orange"]],
                        text="Pct")
        fig_eb.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig_eb.update_xaxes(title="Dependent children")
        fig_eb.update_yaxes(title="")
        apply_wcc_style(fig_eb, "Census 2021, ONS Nomis RM033")
        st.caption("White (37%) and Asian (17%) are the largest ethnic groups")
        st.plotly_chart(fig_eb, use_container_width=True)
        img_btn(fig_eb, "wcc_ethnic")

    with col_e2:
        age_totals = {"0–4":   df_rm006["Age_0_4"].sum(),
                      "5–9":   df_rm006["Age_5_9"].sum(),
                      "10–15": df_rm006["Age_10_15"].sum(),
                      "16–18": df_rm006["Age_16_18"].sum()}
        age_df = pd.DataFrame(list(age_totals.items()), columns=["Age","Count"])
        age_df["Pct"] = age_df["Count"] / age_df["Count"].sum() * 100
        fig_age = px.bar(age_df, x="Age", y="Pct", color="Age",
                         color_discrete_sequence=[WCC["blue"],WCC["cobalt"],WCC["amaranth"],WCC["orange"]],
                         text="Pct")
        fig_age.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig_age.update_yaxes(rangemode="tozero", title="% of dependent children")
        fig_age.update_xaxes(title="Age group")
        apply_wcc_style(fig_age, "Census 2021, ONS Nomis RM006")
        st.caption("0-4 is the largest age group (36%)")
        st.plotly_chart(fig_age, use_container_width=True)
        img_btn(fig_age, "wcc_age")

    st.divider()
    st.subheader("Ethnic diversity vs number of children by LSOA")
    df_sc = df_rm033.copy()
    df_sc["diversity_idx"] = 1 - (
        (df_sc["White"]/df_sc["Total"])**2 + (df_sc["Asian"]/df_sc["Total"])**2 +
        (df_sc["Black"]/df_sc["Total"])**2 + (df_sc["Mixed"]/df_sc["Total"])**2 +
        (df_sc["Other"]/df_sc["Total"])**2)
    df_sc = df_sc.merge(df_rm006[["LSOA_CODE","Total_dep_children"]], on="LSOA_CODE", how="left")
    df_sc = df_sc[df_sc["Total"] > 10]

    fig_sc = px.scatter(df_sc, x="Total", y="diversity_idx", hover_name="LSOA_NAME",
                        size="Total", color="diversity_idx", color_continuous_scale="Blues",
                        labels={"Total":"Total dependent children","diversity_idx":"Diversity index"})
    # Manual numpy trendline (avoids statsmodels dependency)
    _x = df_sc["Total"].values; _y = df_sc["diversity_idx"].values
    _mask = np.isfinite(_x) & np.isfinite(_y)
    if _mask.sum() > 2:
        _m, _b = np.polyfit(_x[_mask], _y[_mask], 1)
        _xr = np.linspace(_x[_mask].min(), _x[_mask].max(), 100)
        fig_sc.add_trace(go.Scatter(x=_xr, y=_m*_xr+_b, mode="lines",
                                    line=dict(color=WCC["amaranth"], width=2, dash="dash"),
                                    name="Trend", showlegend=False))
    apply_wcc_style(fig_sc, "Census 2021, ONS Nomis RM033 — diversity = 1 − Σ(share²)")
    st.caption("Larger LSOAs tend to have higher ethnic diversity")
    st.plotly_chart(fig_sc, use_container_width=True)
    img_btn(fig_sc, "diversity_scatter")

    st.markdown('<div class="source-box">Census 2021, ONS Nomis. RM006: age of youngest dependent child by household type. RM033: ethnic group of dependent child by sex. Both at 2021 LSOA level (123 LSOAs).</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — EGDI
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("Ethnic Group Deprivation Index (EGDI) — Westminster & CIPFA neighbours")
    st.caption("EGDI measures how unevenly deprivation is distributed across ethnic groups within a local authority. Source: Lloyd et al. (2023), gedi.ac.uk/egdi")

    with st.expander("ℹ️ How the EGDI is calculated"):
        st.markdown("""
**The Ethnic Group Deprivation Index (EGDI)** was developed by Lloyd, Catney and colleagues (2023) to measure *within-area* ethnic inequality in deprivation.

**Core metric — EDI score:** For each ethnic group in each LSOA, an Ethnic Deprivation Index (EDI) score is calculated using four deprivation dimensions from the Index of Multiple Deprivation 2019 (IMD):
- **Employment deprivation** (unemployment rate)
- **Education deprivation** (low qualifications rate)
- **Occupational deprivation** (% in routine/semi-routine occupations)
- **Housing deprivation** (% in poor-quality housing or overcrowded)

Each dimension is ranked nationally, standardised, and combined into an overall EDI score per ethnic group per LSOA.

**Range:** The key EGDI summary statistic is the *Range* — the difference between the EDI score of the most deprived and least deprived ethnic group within each LSOA. A higher Range = greater inequality between groups.

**Borough classification:** LAs are classified as having *More*, *Less* or *Similar* ethnic inequality relative to national patterns, based on how their LSOAs distribute across national deprivation deciles.

**Note on NAs:** NAs occur where an ethnic group has fewer than a suppression threshold of individuals in a given LSOA, or is absent entirely.

*Reference: Lloyd, C.D., Catney, G. et al. (2023). The Ethnic Group Deprivation Index. [gedi.ac.uk/egdi](https://gedi.ac.uk/egdi/)*
        """)

    nb_egdi_codes = list(NEIGHBOURS.values())
    df_egdi_nb = df_egdi[df_egdi["LA_Code"].isin(nb_egdi_codes)].copy()
    df_egdi_nb["LA_Name"] = (df_egdi_nb["LA_Name"]
                              .str.replace("and Fulham","& Fulham")
                              .str.replace("and Chelsea","& Chelsea"))

    decile_cols = [f"Pct_D{i}" for i in range(1,11)]
    for c in decile_cols:
        df_egdi_nb[c] = pd.to_numeric(df_egdi_nb[c], errors="coerce")

    df_dec = df_egdi_nb[["LA_Name","Category"]+decile_cols].melt(
        id_vars=["LA_Name","Category"], var_name="Decile", value_name="Pct")
    df_dec["Decile_n"] = df_dec["Decile"].str.replace("Pct_D","").astype(float)
    df_dec = df_dec.dropna(subset=["Pct"])

    col_g1, col_g2 = st.columns([3,2])
    with col_g1:
        fig_dec = px.line(df_dec.sort_values("Decile_n"), x="Decile_n", y="Pct",
                          color="LA_Name", markers=True,
                          color_discrete_map={la: BOROUGH_COLOURS.get(la, WCC["blue"])
                                             for la in df_dec["LA_Name"].unique()},
                          labels={"Decile_n":"Income deprivation decile (1 = most deprived)",
                                  "Pct":"% of borough LSOAs", "LA_Name":""})
        fig_dec.update_traces(line_width=2.5)
        fig_dec.update_xaxes(tickvals=list(range(1,11)))
        fig_dec.update_yaxes(rangemode="tozero")
        fig_dec.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                        font=dict(size=10, family="Arial"), title=None),
            margin=dict(l=50, r=20, t=50, b=80),
            height=380,
            font_family="Arial", font_color=WCC["blue"],
            plot_bgcolor="white", paper_bgcolor="white",
        )
        fig_dec.update_xaxes(showgrid=False, linecolor="#ccc", showline=True, tickfont=dict(size=11))
        fig_dec.update_yaxes(gridcolor=WCC["grid"], zeroline=False, tickfont=dict(size=11))
        fig_dec.add_annotation(
            text="<i>Source: EGDI (Lloyd et al. 2023) — gedi.ac.uk/egdi. Based on IMD 2019.</i>",
            xref="paper", yref="paper", x=0, y=-0.28,
            showarrow=False, font=dict(size=9, color="#666", family="Arial"), align="left")
        st.caption("Westminster: 22% of LSOAs fall in the most deprived national decile")
        st.plotly_chart(fig_dec, use_container_width=True)
        img_btn(fig_dec, "egdi_decile")

    with col_g2:
        st.markdown("**EGDI Classification**")
        for _, row in df_egdi_nb.iterrows():
            cat = str(row.get("Category",""))
            la  = row["LA_Name"]
            col = BOROUGH_COLOURS.get(la, WCC["blue"])
            icon = "🔴" if "More" in cat else "🟢" if "Less" in cat else "⚪"
            st.markdown(
                f"<div style='border-left:3px solid {col};padding:6px 10px;margin-bottom:6px;"
                f"background:{WCC['light_blue']};border-radius:4px'>"
                f"<b>{la}</b> {icon}<br><small>{cat}</small></div>",
                unsafe_allow_html=True)
        st.caption("🔴 More ethnic inequality = deprivation is distributed more unequally across ethnic groups. Does not mean the borough is more deprived overall.")

    # ── LSOA-level EGDI visualisations
    if not df_egdi_lsoa.empty:
        st.divider()
        st.subheader("Westminster LSOA-level Ethnic Deprivation (EGDI)")
        st.caption("EDI scores by ethnic group and LSOA. NAs = ethnic group absent or below threshold in that LSOA.")

        egdi_tabs = st.tabs(["📊 Range by LSOA","🗺️ EGDI Map","🔥 Ethnic heatmap"])

        with egdi_tabs[0]:
            # Range = difference between most and least deprived ethnic group EDI score
            df_range = df_egdi_lsoa[["LSOA21CD","LSOA21NM","Range",
                                     "Mostdeprivedgroup","Leastdeprivedgroup"]].dropna(subset=["Range"])
            df_range = df_range.sort_values("Range", ascending=False)
            fig_range = px.bar(df_range.head(30).sort_values("Range"),
                               x="Range", y="LSOA21NM", orientation="h",
                               color="Range", color_continuous_scale="RdYlGn_r",
                               hover_data={"Mostdeprivedgroup":True,"Leastdeprivedgroup":True},
                               labels={"Range":"EGDI Range","LSOA21NM":"LSOA"})
            fig_range.update_xaxes(title="EGDI Range (higher = more inequality between ethnic groups)")
            fig_range.update_yaxes(title="")
            fig_range.update_layout(coloraxis_showscale=False, margin=dict(l=150,r=20,t=30,b=50),
                                    height=500, font_family="Arial", plot_bgcolor="white",
                                    paper_bgcolor="white")
            fig_range.update_xaxes(showgrid=False)
            fig_range.update_yaxes(tickfont=dict(size=9))
            fig_range.add_annotation(
                text="<i>Source: EGDI (Lloyd et al. 2023). Range = EDI score of most deprived group minus least deprived group within each LSOA.</i>",
                xref="paper", yref="paper", x=0, y=-0.09,
                showarrow=False, font=dict(size=9, color="#666", family="Arial"))
            st.caption("Top 30 Westminster LSOAs by EGDI Range — highest ethnic inequality")
            st.plotly_chart(fig_range, use_container_width=True)
            img_btn(fig_range, "egdi_range")

        with egdi_tabs[1]:
            # Map EGDI Range across LSOAs
            map_lsoa = df_egdi_lsoa[["LSOA21CD","LSOA21NM","Range"]].dropna(subset=["Range"])
            lsoa_codes_m = map_lsoa["LSOA21CD"].tolist()
            lsoa_z_m     = map_lsoa["Range"].tolist()
            lsoa_names_m = map_lsoa["LSOA21NM"].tolist()

            fig_egdi_map = go.Figure(go.Choroplethmap(
                geojson=wcc_geojson,
                locations=lsoa_codes_m,
                z=lsoa_z_m,
                text=lsoa_names_m,
                hovertemplate="<b>%{text}</b><br>EGDI Range: %{z:.3f}<extra></extra>",
                colorscale="RdYlGn_r",
                marker_opacity=0.8, marker_line_width=0.3,
                colorbar=dict(title=dict(text="EGDI Range"), thickness=14, len=0.6),
            ))
            fig_egdi_map.update_layout(
                map_style="carto-positron", map_zoom=12,
                map_center={"lat":51.512,"lon":-0.155},
                margin=dict(l=0,r=0,t=0,b=0), height=500,
                paper_bgcolor="white",
            )
            st.caption("EGDI Range by Westminster LSOA — red = higher inequality between ethnic groups")
            st.plotly_chart(fig_egdi_map, use_container_width=True)
            img_btn(fig_egdi_map, "egdi_lsoa_map")

        with egdi_tabs[2]:
            # Heatmap: ethnic groups vs LSOAs (EDI scores)
            edi_cols = [c for c in df_egdi_lsoa.columns if c.startswith("EDI.")]
            # Rename cols for readability
            rename = {c: c.replace("EDI.","") for c in edi_cols}
            hm_df = (df_egdi_lsoa[["LSOA21NM"] + edi_cols]
                     .rename(columns=rename)
                     .set_index("LSOA21NM"))
            # Keep groups with >30% non-NA
            eth_keep = [c for c in hm_df.columns if hm_df[c].notna().mean() > 0.3]
            hm_df = hm_df[eth_keep]

            fig_hm = go.Figure(go.Heatmap(
                z=hm_df.values,
                x=hm_df.columns.tolist(),
                y=hm_df.index.tolist(),
                colorscale="RdYlGn_r",
                hoverongaps=False,
                hovertemplate="LSOA: %{y}<br>Group: %{x}<br>EDI: %{z:.2f}<extra></extra>",
                colorbar=dict(title="EDI score", thickness=14),
            ))
            fig_hm.update_layout(
                height=700,
                margin=dict(l=130, r=20, t=50, b=100),
                font_family="Arial", font_color=WCC["blue"],
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(tickangle=-45, tickfont=dict(size=10)),
                yaxis=dict(tickfont=dict(size=8), autorange="reversed"),
            )
            fig_hm.add_annotation(
                text="<i>NAs (grey) = ethnic group absent or below threshold in that LSOA. Source: EGDI (Lloyd et al. 2023).</i>",
                xref="paper", yref="paper", x=0, y=-0.13,
                showarrow=False, font=dict(size=9, color="#666", family="Arial"))
            st.caption("EDI scores by ethnic group and LSOA — red = higher deprivation, green = lower")
            st.plotly_chart(fig_hm, use_container_width=True)
            img_btn(fig_hm, "egdi_heatmap")

    # ── LA-level radar
    st.divider()
    st.subheader("Deprivation profile — CIPFA neighbours")
    radar_cols = [f"Pct_D{i}" for i in range(1,6)]
    radar_labels = ["Decile 1\n(most deprived)","Decile 2","Decile 3","Decile 4","Decile 5"]
    for c in radar_cols:
        df_egdi_nb[c] = pd.to_numeric(df_egdi_nb[c], errors="coerce") * 100
    fig_r = go.Figure()
    for _, row in df_egdi_nb.iterrows():
        la   = row["LA_Name"]
        vals = [float(row[c]) if not pd.isna(row[c]) else 0 for c in radar_cols]
        fig_r.add_trace(go.Scatterpolar(
            r=vals+[vals[0]], theta=radar_labels+[radar_labels[0]],
            name=la, line=dict(color=BOROUGH_COLOURS.get(la, WCC["blue"]), width=2),
            fill="toself", fillcolor=BOROUGH_COLOURS.get(la, WCC["blue"]), opacity=0.12))
    fig_r.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0,30])),
        font_family="Arial", paper_bgcolor="white", height=420,
        legend=dict(orientation="h", y=-0.15, x=0))
    st.caption("Westminster and Islington: most LSOAs in most deprived deciles")
    st.plotly_chart(fig_r, use_container_width=True)
    img_btn(fig_r, "egdi_radar")

    st.markdown('<div class="source-box">EGDI (Ethnic Group Deprivation Index), Lloyd et al. (2023). Based on IMD 2019 income deprivation domain. "More ethnic inequality" means deprivation is significantly unequally distributed across ethnic groups within the borough — not that it is more deprived overall. NAs in LSOA data indicate LSOAs where there is no one in a given ethnic group, or the group count is below threshold.</div>', unsafe_allow_html=True)
