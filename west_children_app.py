# Westminster Children's Demographics Dashboard
# Run: streamlit run west_children_app.py
#
# DATA SOURCES
# 1. DWP Children in Low Income Families (AHC Relative) FYE 2025
#    https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025
# 2. DfE Key Stage 4 attainment by ethnicity — Explore Education Statistics
#    https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df
# 3. EGDI Ethnic Group Deprivation Index — Lloyd et al. 2023 / gedi.ac.uk
# 4. Census 2021 RM006, RM033 — ONS Nomis
#    https://www.nomisweb.co.uk/
# 5. Westminster LSOA boundaries — ONS Open Geography Portal (2021 boundaries)
# 6. London borough boundaries — London Datastore

import os, io, json, warnings
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import topojson as tp
from pptx import Presentation
from pptx.util import Inches, Pt

warnings.filterwarnings("ignore")

# ── DATA PATHS ────────────────────────────────────────────────────────────────
_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_DATA_SUBDIR = os.path.join(_SCRIPT_DIR, "data")
_SEARCH_DIRS = [_DATA_SUBDIR, _SCRIPT_DIR, "/mnt/user-data/uploads"]

_ALIASES = {
    "data-key-stage-4-performance__1_.csv": [
        "data-key-stage-4-performance__1_.csv",
        "data-key-stage-4-performance_1_.csv",
        "data-key-stage-4-performance (1).csv",
    ],
    "ONS_LSOA_2021 (1).json": [
        "ONS_LSOA_2021 (1).json",
        "ONS_LSOA_2021__1_.json",
        "ONS_LSOA_2021_(1).json",
        "ONS_LSOA_2021 (1) .json",
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

# ── WCC BRAND COLOURS (from official brand guidelines) ───────────────────────
WCC = {
    "blue":    "#0B2265",   # Westminster Blue — primary (Pantone 2758C)
    "yellow":  "#F5CB00",   # Yellow — primary (Pantone 116C)
    "cobalt":  "#0C35FA",   # Cobalt — secondary
    "amaranth":"#E34063",   # Amaranth — secondary (Pantone 7424C)
    "green":   "#008466",   # Green (from accessibility chart)
    "orange":  "#EA6F06",   # Orange (from accessibility chart)
    "white":   "#FFFFFF",
    "text":    "#0B2265",   # Body text: Westminster Blue
    "grid":    "#E8E8E8",
    "bg":      "#FFFFFF",
    "light_blue": "#E8EBF5",  # 10% tint of Westminster Blue
}

# Borough colours using WCC palette
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
st.set_page_config(page_title="Westminster Children's Demographics",
                   page_icon="🏙️", layout="wide", initial_sidebar_state="expanded")

st.markdown(f"""<style>
  /* WCC brand colours */
  :root {{
    --wcc-blue:    {WCC["blue"]};
    --wcc-yellow:  {WCC["yellow"]};
    --wcc-cobalt:  {WCC["cobalt"]};
    --wcc-amaranth:{WCC["amaranth"]};
    --wcc-light:   {WCC["light_blue"]};
  }}
  /* Sidebar: Westminster Blue background with yellow keyline at top */
  [data-testid="stSidebar"] {{
    background-color: {WCC["blue"]} !important;
    border-top: 4px solid {WCC["yellow"]};
  }}
  [data-testid="stSidebar"] * {{ color: #ffffff !important; }}
  [data-testid="stSidebar"] a {{ color: #A8C0FF !important; }}
  [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.25) !important; }}
  /* Headings */
  h1, h2, h3 {{ color: {WCC["blue"]} !important; font-family: Arial, sans-serif !important; }}
  /* Tabs */
  .stTabs [data-baseweb="tab"] {{ font-size: .95rem; font-family: Arial, sans-serif; }}
  .stTabs [aria-selected="true"] {{
    color: {WCC["blue"]} !important;
    border-bottom: 3px solid {WCC["yellow"]} !important;
    font-weight: 700;
  }}
  /* Metric cards */
  [data-testid="stMetric"] {{
    background: {WCC["light_blue"]};
    border-radius: 6px;
    padding: 12px 14px;
    border-left: 4px solid {WCC["blue"]};
  }}
  /* Info/alert boxes */
  .stAlert {{ border-left: 4px solid {WCC["yellow"]} !important; }}
  /* Source box */
  .source-box {{
    background: {WCC["light_blue"]};
    border-radius: 5px;
    padding: 8px 12px;
    font-size: .82em;
    color: #333;
    margin-top: 6px;
    border-left: 3px solid {WCC["blue"]};
  }}
  /* Reduce chart title crowding */
  .js-plotly-plot .plotly .gtitle {{ font-size: 13px !important; }}
</style>""", unsafe_allow_html=True)

# ── HELPERS ───────────────────────────────────────────────────────────────────
def _pct(s):
    v = pd.to_numeric(s.astype(str).str.replace("%","",regex=False)
                       .str.replace(",","",regex=False).str.strip(), errors="coerce")
    if v.dropna().max() <= 1.0: v = v * 100
    return v

def _num(s):
    return pd.to_numeric(s.astype(str).str.replace(",","",regex=False).str.strip(), errors="coerce")

# ── DATA LOADERS ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_low_income_la():
    df = pd.read_csv(_dp("2_AHC_Relative_LA.csv"), header=8,
                     names=["LA","Area_Code","N_2024","N_2025","Pct_2024","Pct_2025"],
                     usecols=[0,1,2,3,4,5])
    df = df.dropna(subset=["Area_Code"])
    df["Pct_2024"] = _pct(df["Pct_2024"])
    df["Pct_2025"] = _pct(df["Pct_2025"])
    df["N_2024"] = _num(df["N_2024"])
    df["N_2025"] = _num(df["N_2025"])
    return df

@st.cache_data(show_spinner=False)
def load_low_income_ward():
    df = pd.read_csv(_dp("4_AHC_Relative_Ward.csv"), header=9,
                     names=["LA","LA_Code","Ward","Ward_Code","N_2024","N_2025","Pct_2024","Pct_2025"],
                     usecols=[0,1,2,3,4,5,6,7])
    df = df.dropna(subset=["Ward_Code"])
    df["Pct_2024"] = _pct(df["Pct_2024"])
    df["Pct_2025"] = _pct(df["Pct_2025"])
    df["N_2024"] = _num(df["N_2024"])
    df["N_2025"] = _num(df["N_2025"])
    df["LA_filled"] = df["LA"].ffill()
    return df

@st.cache_data(show_spinner=False)
def load_ks4_ethnic():
    """KS4 Attainment 8 by ethnic group — Inner London LAs (2024/25)."""
    path = _dp("data-key-stage-4-performance__1_.csv")
    df = pd.read_csv(path, header=None, low_memory=False)

    years_row   = df.iloc[2, 4:].ffill()
    metrics_row = df.iloc[3, 4:].ffill()
    sub_row     = df.iloc[4, 4:]

    # Build unique column names — avoids narwhals DuplicateError
    cols = ["ethnic_group", "subgroup", "region", "la"]
    seen = {}
    for j in range(len(years_row)):
        yr = str(years_row.iloc[j]).strip().replace("/","_")
        m  = str(metrics_row.iloc[j]).strip().replace(" ","_")[:22]
        s  = str(sub_row.iloc[j]).strip().replace(" ","_")[:18]
        base = f"{yr}_{m}_{s}"
        seen[base] = seen.get(base, 0) + 1
        cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")

    data = df.iloc[5:].copy()
    data.columns = cols[:len(data.columns)]
    data["ethnic_group"] = data["ethnic_group"].ffill()
    data["subgroup"]     = data["subgroup"].ffill()
    data["region"]       = data["region"].ffill()

    inner = ["Camden","Hackney","Hammersmith and Fulham","Haringey","Islington",
             "Kensington and Chelsea","Lambeth","Lewisham","Newham","Southwark",
             "Tower Hamlets","Wandsworth","Westminster"]

    mask = data["la"].isin(inner) & data["subgroup"].astype(str).str.startswith("All")
    sub  = data[mask].copy()

    att_col = [c for c in sub.columns if "2024_25" in c and "Attainment_8" in c and "Total" in c]
    pct_col = [c for c in sub.columns if "2024_25" in c and "achieving_gr" in c and "Total" in c]

    sub["att8_2425"] = pd.to_numeric(sub[att_col[0]], errors="coerce") if att_col else np.nan
    if pct_col:
        sub["pct_5above_EM_2425"] = pd.to_numeric(
            sub[pct_col[0]].astype(str).str.replace("%","",regex=False), errors="coerce")
    else:
        sub["pct_5above_EM_2425"] = np.nan

    sub = sub.copy()
    sub["la"] = (sub["la"].astype(str)
                 .str.replace("and Fulham","& Fulham",regex=False)
                 .str.replace("and Chelsea","& Chelsea",regex=False))
    return sub

@st.cache_data(show_spinner=False)
def load_ks4_time():
    """KS4 time series (all pupils) — Inner London LAs. Returns empty DF if file missing."""
    path = _dp("data-key-stage-4-performance.csv")
    if not os.path.exists(path):
        return pd.DataFrame(columns=["la","ethnic_group","subgroup","sex","year","att8"])

    df = pd.read_csv(path, header=None, low_memory=False)
    YEARS = ["2018/19","2019/20","2020/21","2021/22","2022/23","2023/24","2024/25"]

    metrics_row = df.iloc[2, 5:].astype(str).fillna("").values
    att8_start  = next((j+5 for j,v in enumerate(metrics_row) if "Attainment 8" in str(v)), 12)

    df_c = df.copy()
    for ci in [0,1,2,4]:
        if ci < df_c.shape[1]:
            df_c.iloc[4:, ci] = df_c.iloc[4:, ci].ffill()

    inner = {"Camden","Hackney","Hammersmith and Fulham","Haringey","Islington",
             "Kensington and Chelsea","Lambeth","Lewisham","Newham","Southwark",
             "Tower Hamlets","Wandsworth","Westminster"}

    records = []
    for i in range(4, len(df_c)):
        la = str(df_c.iloc[i, 4]).strip() if 4 < df_c.shape[1] else None
        if not la or la not in inner:
            continue
        eg  = str(df_c.iloc[i, 0]).strip() if pd.notna(df_c.iloc[i, 0]) else "All pupils"
        sg  = str(df_c.iloc[i, 1]).strip() if pd.notna(df_c.iloc[i, 1]) else ""
        sex = str(df_c.iloc[i, 2]).strip() if pd.notna(df_c.iloc[i, 2]) else ""
        for yr_idx, yr in enumerate(YEARS):
            cidx = att8_start + yr_idx
            if cidx < df_c.shape[1]:
                val = pd.to_numeric(str(df_c.iloc[i,cidx]).replace(",","").strip(), errors="coerce")
                records.append({"la":la,"ethnic_group":eg,"subgroup":sg,"sex":sex,"year":yr,"att8":val})

    ts = pd.DataFrame(records) if records else pd.DataFrame(
        columns=["la","ethnic_group","subgroup","sex","year","att8"])
    if not ts.empty:
        ts["la"] = (ts["la"].str.replace("and Fulham","& Fulham",regex=False)
                             .str.replace("and Chelsea","& Chelsea",regex=False))
    return ts

@st.cache_data(show_spinner=False)
def load_rm006():
    df = pd.read_excel(_dp("RM006_age_of_youngest_dependent_child_by_household_type.xlsx"),
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
    for c in df.columns[1:-2]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    eth = [c for c in df.columns if c not in ["LSOA","LSOA_CODE","LSOA_NAME"]]
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
    """Westminster 2021 LSOA boundaries — standard GeoJSON, 123 LSOAs."""
    with open(_dp("ONS_LSOA_2021 (1).json")) as f:
        return json.load(f)

@st.cache_data(show_spinner=False)
def load_borough_geojson():
    """London borough boundaries from TopoJSON — for CIPFA neighbour choropleth."""
    with open(_dp("Borough_London_LL84.json")) as f:
        topo = json.load(f)
    return json.loads(tp.Topology(topo, object_name="Borough_London_LL84").to_geojson())

# ── PPTX HELPERS ──────────────────────────────────────────────────────────────
def _fig_to_png(fig):
    """Export Plotly figure to PNG bytes. Falls back gracefully if kaleido absent."""
    try:
        return fig.to_image(format="png", width=1200, height=680, scale=2)
    except Exception:
        # kaleido/Chrome not available — return a minimal valid PNG
        import base64
        return base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")

def _fig_to_pptx(fig, title=""):
    img = _fig_to_png(fig)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(io.BytesIO(img), Inches(0.3), Inches(0.3), Inches(12.73), Inches(6.4))
    txb = slide.shapes.add_textbox(Inches(0.3), Inches(6.8), Inches(12), Inches(0.55))
    tf = txb.text_frame
    tf.text = title or (fig.layout.title.text or "Chart")
    tf.paragraphs[0].runs[0].font.size = Pt(13)
    tf.paragraphs[0].runs[0].font.bold = True
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf

def pptx_btn(fig, key, title=""):
    try:
        buf = _fig_to_pptx(fig, title)
        st.download_button("⬇ Download slide (PPTX)", buf, f"{key}.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key=f"dl_{key}")
    except Exception as e:
        st.caption(f"_Slide export unavailable ({e})_")

def apply_wcc_style(fig, source=""):
    """Apply Westminster City Council brand styling to a Plotly figure."""
    fig.update_layout(
        font_family="Arial",
        font_color=WCC["blue"],
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font_size=13,          # Smaller to avoid crowding
        title_font_color=WCC["blue"],
        title_font_family="Arial",
        margin=dict(l=50, r=30, t=60, b=55),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                    font=dict(family="Arial", size=11)),
    )
    fig.update_xaxes(showgrid=False, linecolor="#cccccc", showline=True,
                     tickfont=dict(family="Arial", size=11))
    fig.update_yaxes(gridcolor=WCC["grid"], linecolor="white", zeroline=False,
                     tickfont=dict(family="Arial", size=11))
    if source:
        fig.add_annotation(
            text=f"<i>Source: {source}</i>",
            xref="paper", yref="paper", x=0, y=-0.14,
            showarrow=False, font=dict(size=10, color="#555", family="Arial"), align="left")
    return fig

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    logo_path = _dp("city_of_westminster.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=160)
    else:
        st.image("https://raw.githubusercontent.com/SianReilly/childrens_demographics_wcc/main/data/city_of_westminster.png",
                 width=160)
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
    wcc_geojson    = load_wcc_geojson()
    borough_geojson = load_borough_geojson()

# ── HEADER & METRICS ──────────────────────────────────────────────────────────
st.title("🏙️ Westminster Children's Demographics")
st.markdown("Child poverty, demographics, attainment and ethnic diversity — benchmarked against CIPFA statistical neighbours.")

wcc_li_rows = df_li_la[df_li_la["LA"].str.contains("Westminster", na=False)]
wcc_li = wcc_li_rows.iloc[0] if len(wcc_li_rows) else pd.Series({"N_2024":0,"N_2025":0,"Pct_2024":0,"Pct_2025":0})

london_avg  = df_li_la[df_li_la["Area_Code"].astype(str).str.startswith("E09",na=False)]["Pct_2025"].mean()
total_dep   = df_rm006["Total_dep_children"].sum()
pct_u5      = round(df_rm006["Age_0_4"].sum()/total_dep*100, 1) if total_dep else 0
rm_tot      = df_rm033["Total"].sum()
pct_nonwhite = round((1 - df_rm033["White"].sum()/rm_tot)*100, 1) if rm_tot else 0

c1,c2,c3,c4,c5 = st.columns(5)
c1.metric(
    "Children in low income (FYE 2025)",
    f"{int(wcc_li['N_2025']):,}",
    delta=f"↕ {wcc_li['Pct_2025']:.1f}% of all children in Westminster",
    delta_color="off",
    help="Number of children aged 0–15 in families in relative low income after housing costs (AHC), FYE 2025. Source: DWP."
)
c2.metric(
    "Change vs FYE 2024",
    f"{int(wcc_li['N_2025']-wcc_li['N_2024']):+,} children",
    delta=f"{wcc_li['Pct_2025']-wcc_li['Pct_2024']:+.1f} percentage points",
    delta_color="inverse",
    help="Change in the number and rate of children in low income families between FYE 2024 and FYE 2025. A negative change (shown in green) means fewer children in poverty."
)
c3.metric(
    "Westminster vs London average",
    f"{wcc_li['Pct_2025']:.1f}% (Westminster)",
    delta=f"{wcc_li['Pct_2025']-london_avg:+.1f}pp vs London borough average ({london_avg:.1f}%)",
    delta_color="inverse",
    help="Westminster's child poverty rate (AHC relative) compared to the unweighted average across all 33 London boroughs. A negative value means Westminster is below the London average."
)
c4.metric(
    "Dependent children (Census 2021)",
    f"{total_dep:,}",
    delta=f"{pct_u5}% aged 0–4 (youngest age group)",
    delta_color="off",
    help="Total dependent children across 123 Westminster LSOAs. Census 2021, ONS Nomis RM006."
)
c5.metric(
    "Non-white children (Census 2021)",
    f"{pct_nonwhite}%",
    delta="of dependent children across Westminster LSOAs",
    delta_color="off",
    help="Percentage of dependent children who are not classified as White, based on the ethnic group of the dependent child. Census 2021, ONS Nomis RM033."
)
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
    df_nb["Borough"] = (df_nb["LA"].str.replace("and Fulham","& Fulham")
                                    .str.replace("and Chelsea","& Chelsea"))
    df_nb = df_nb.sort_values("Pct_2025", ascending=True)

    col_a, col_b = st.columns([3,2])
    with col_a:
        fig1 = px.bar(df_nb, x="Pct_2025", y="Borough", orientation="h",
                      title="Islington highest; Kensington & Chelsea lowest",
                      color="Borough",
                      color_discrete_map={r["Borough"]: BOROUGH_COLOURS.get(r["Borough"], WCC["blue"])
                                          for _, r in df_nb.iterrows()},
                      text="Pct_2025")
        fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig1.update_xaxes(range=[0, df_nb["Pct_2025"].max()*1.22], title="% children in low income (AHC)")
        fig1.update_yaxes(title="")
        apply_wcc_style(fig1, "DWP Children in Low Income Families FYE 2025")
        st.plotly_chart(fig1, use_container_width=True)
        pptx_btn(fig1, "child_poverty_bar", "Child poverty rate — CIPFA neighbours FYE 2025")

    with col_b:
        df_ts = df_nb[["Borough","Pct_2024","Pct_2025"]].melt("Borough",var_name="Year",value_name="Pct")
        df_ts["Year"] = df_ts["Year"].map({"Pct_2024":"FYE 2024","Pct_2025":"FYE 2025"})
        fig2 = px.line(df_ts, x="Year", y="Pct", color="Borough", markers=True,
                       title="All neighbours: poverty fell 2024→2025",
                       color_discrete_map={b:BOROUGH_COLOURS.get(b,WCC["blue"]) for b in df_ts["Borough"].unique()})
        fig2.update_traces(line_width=2.5)
        fig2.update_yaxes(title="% children in low income", rangemode="tozero")
        fig2.update_xaxes(title="")
        apply_wcc_style(fig2, "DWP Children in Low Income Families FYE 2025")
        st.plotly_chart(fig2, use_container_width=True)
        pptx_btn(fig2, "child_poverty_trend", "Child poverty trend 2024–2025")

    # CIPFA neighbours choropleth map
    st.divider()
    st.subheader("Geographic context — CIPFA statistical neighbours")
    nb_gj = {"type":"FeatureCollection",
              "features":[f for f in borough_geojson["features"]
                          if f["properties"]["BoroughCod"] in nb_codes]}
    fig_nb_map = px.choropleth_map(
        df_nb, geojson=nb_gj, locations="Area_Code",
        featureidkey="properties.BoroughCod",
        color="Pct_2025", hover_name="Borough",
        hover_data={"Pct_2025":":.1f%","Area_Code":False},
        color_continuous_scale=[[0,"#E8EBF5"],[0.5,"#0C35FA"],[1.0,"#0B2265"]],
        range_color=[df_nb["Pct_2025"].min()*0.9, df_nb["Pct_2025"].max()*1.05],
        zoom=10.5, center={"lat":51.505,"lon":-0.17}, opacity=0.75,
        map_style="carto-positron",
        title="% children in low income — CIPFA neighbours (FYE 2025)",
        labels={"Pct_2025":"% in low income"})
    fig_nb_map.update_layout(height=460, margin=dict(l=0,r=0,t=50,b=0))
    st.plotly_chart(fig_nb_map, use_container_width=True)
    pptx_btn(fig_nb_map, "cipfa_map", "CIPFA neighbours child poverty map FYE 2025")

    st.divider()
    st.subheader("Westminster ward-level child poverty (FYE 2025)")
    st.caption("Source: DWP/HMRC Children in Low Income Families, AHC Relative, ward level.")
    wcc_wards = df_li_ward[df_li_ward["LA_filled"].astype(str).str.contains("Westminster",na=False)].copy()
    wcc_wards = wcc_wards.sort_values("Pct_2025",ascending=True).dropna(subset=["Ward","Pct_2025"])
    fig3 = px.bar(wcc_wards, x="Pct_2025", y="Ward", orientation="h",
                  title="Church Street & Westbourne have highest child poverty",
                  color_discrete_sequence=[WCC["blue"]], text="Pct_2025")
    fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig3.update_xaxes(range=[0,wcc_wards["Pct_2025"].max()*1.22], title="% children in low income (AHC)")
    fig3.update_yaxes(title="")
    apply_wcc_style(fig3, "DWP Children in Low Income Families FYE 2025 — ward level")
    st.plotly_chart(fig3, use_container_width=True)
    pptx_btn(fig3, "ward_poverty", "Westminster ward child poverty FYE 2025")

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
        Total_dep_children=("Total_dep_children","sum"),
        Age_0_4=("Age_0_4","sum"), Age_10_15=("Age_10_15","sum"),
        White=("White","sum"), Asian=("Asian","sum"),
        Black=("Black","sum"), Mixed=("Mixed","sum"), Total=("Total","sum")
    ).reset_index()

    sd = lambda n,d: np.where(d>0, n/d*100, 0)
    agg["pct_u5"]    = sd(agg["Age_0_4"],   agg["Total_dep_children"])
    agg["pct_10_15"] = sd(agg["Age_10_15"], agg["Total_dep_children"])
    agg["pct_white"] = sd(agg["White"],     agg["Total"])
    agg["pct_asian"] = sd(agg["Asian"],     agg["Total"])
    agg["pct_black"] = sd(agg["Black"],     agg["Total"])
    agg["pct_mixed"] = sd(agg["Mixed"],     agg["Total"])

    METRIC_MAP = {
        "Total dependent children": ("Total_dep_children","Total dependent children","Blues"),
        "% aged 0–4":               ("pct_u5",   "% aged 0–4",      "YlOrBr"),
        "% aged 10–15":             ("pct_10_15","% aged 10–15",     "Purples"),
        "% White children":         ("pct_white","% White children", "Greys"),
        "% Asian children":         ("pct_asian","% Asian children", "Oranges"),
        "% Black children":         ("pct_black","% Black children", "Reds"),
        "% Mixed children":         ("pct_mixed","% Mixed children", "Greens"),
    }
    col_name, label, cscale = METRIC_MAP[map_metric]

    fig_map = px.choropleth_map(
        agg, geojson=wcc_geojson, locations="LSOA_CODE",
        featureidkey="properties.LSOA21CD",
        color=col_name, hover_name="LSOA_NAME",
        hover_data={col_name:":.1f","LSOA_CODE":True},
        color_continuous_scale=cscale, zoom=12,
        center={"lat":51.512,"lon":-0.155}, opacity=0.75,
        title=f"{label} by Westminster LSOA (Census 2021)",
        map_style="carto-positron", labels={col_name:label})
    fig_map.update_layout(margin=dict(l=0,r=0,t=50,b=0), height=560)
    st.plotly_chart(fig_map, use_container_width=True)
    pptx_btn(fig_map, f"lsoa_{col_name}", f"{label} — Westminster LSOAs (Census 2021)")

    st.markdown('<div class="source-box">Census 2021, ONS Nomis. RM006: household type by age of youngest child. RM033: ethnic group of dependent child by sex. ONS 2021 LSOA boundaries.</div>', unsafe_allow_html=True)

    st.divider()
    top15 = agg.nlargest(15,"Total_dep_children").sort_values("Total_dep_children")
    fig_lsoa = px.bar(top15, x="Total_dep_children", y="LSOA_NAME", orientation="h",
                      title="Top 15 LSOAs: most dependent children",
                      color_discrete_sequence=[WCC["blue"]], text="Total_dep_children")
    fig_lsoa.update_traces(textposition="outside")
    fig_lsoa.update_xaxes(title="Total dependent children (Census 2021)")
    fig_lsoa.update_yaxes(title="")
    apply_wcc_style(fig_lsoa, "Census 2021, ONS Nomis RM006")
    st.plotly_chart(fig_lsoa, use_container_width=True)
    pptx_btn(fig_lsoa, "lsoa_top15", "Top 15 Westminster LSOAs by dependent children")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — KS4 ATTAINMENT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Key Stage 4 attainment — Westminster vs CIPFA neighbours")
    st.caption("Average Attainment 8 score, 2024/25, state-funded schools. Source: DfE Explore Education Statistics.")

    nb_la = ["Camden","Hammersmith & Fulham","Islington","Kensington & Chelsea","Wandsworth","Westminster"]
    ks4_nb = df_ks4_eth[df_ks4_eth["la"].isin(nb_la) & df_ks4_eth["att8_2425"].notna()].copy()

    col_ka, col_kb = st.columns(2)
    with col_ka:
        wcc_eth = ks4_nb[ks4_nb["la"]=="Westminster"].dropna(subset=["att8_2425"])
        if not wcc_eth.empty:
            # Rename columns to be unique before plotting
            plot_df = wcc_eth[["ethnic_group","att8_2425"]].copy()
            plot_df.columns = ["ethnic_group","att8_2425"]
            plot_df = plot_df.drop_duplicates("ethnic_group")
            fig_ae = px.bar(plot_df.sort_values("att8_2425"),
                            x="att8_2425", y="ethnic_group", orientation="h",
                            title="Westminster: Attainment 8 by ethnic group (2024/25)",
                            color_discrete_sequence=[WCC["blue"]], text="att8_2425")
            fig_ae.update_traces(texttemplate="%{text:.1f}", textposition="outside")
            fig_ae.update_xaxes(title="Average Attainment 8 score", range=[0,75])
            fig_ae.update_yaxes(title="")
            apply_wcc_style(fig_ae, "DfE KS4 2024/25, state-funded schools")
            st.plotly_chart(fig_ae, use_container_width=True)
            pptx_btn(fig_ae, "wcc_att8_ethnic", "Westminster Attainment 8 by ethnic group 2024/25")

    with col_kb:
        all_eth = ks4_nb[["la","ethnic_group","att8_2425"]].dropna(subset=["att8_2425"]).copy()
        all_eth = all_eth.drop_duplicates(["la","ethnic_group"])
        if not all_eth.empty:
            fig_ac = px.bar(all_eth.sort_values(["ethnic_group","att8_2425"]),
                            x="att8_2425", y="la", orientation="h",
                            facet_col="ethnic_group", facet_col_wrap=2,
                            title="Attainment 8 by ethnic group and borough",
                            color="la",
                            color_discrete_map={la:BOROUGH_COLOURS.get(la,WCC["blue"]) for la in all_eth["la"].unique()})
            fig_ac.update_traces(showlegend=False)
            fig_ac.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1][:25]))
            fig_ac.update_xaxes(range=[0,80], matches=None)
            fig_ac.update_yaxes(matches=None)
            fig_ac.update_layout(height=480)
            apply_wcc_style(fig_ac, "DfE KS4 2024/25")
            st.plotly_chart(fig_ac, use_container_width=True)
            pptx_btn(fig_ac, "att8_boroughs", "Attainment 8 by borough and ethnic group")

    # Time series — only show if data loaded
    st.divider()
    st.subheader("Attainment 8 trend 2018/19–2024/25")
    if df_ks4_ts.empty or "la" not in df_ks4_ts.columns:
        st.info("Time series data not available — upload `data-key-stage-4-performance.csv` to the `data/` folder in GitHub.")
    else:
        nb_v1 = ["Camden","Hammersmith and Fulham","Islington","Kensington and Chelsea","Wandsworth","Westminster"]
        ts_d = df_ks4_ts[
            df_ks4_ts["la"].isin(nb_v1) &
            df_ks4_ts["subgroup"].str.startswith("All", na=False) &
            df_ks4_ts["att8"].notna()
        ].copy()
        ts_d["la"] = (ts_d["la"].str.replace("and Fulham","& Fulham",regex=False)
                                 .str.replace("and Chelsea","& Chelsea",regex=False))
        ts_agg = ts_d.groupby(["la","year"])["att8"].mean().reset_index()
        if not ts_agg.empty:
            fig_ts = px.line(ts_agg, x="year", y="att8", color="la", markers=True,
                             title="Attainment 8 dipped 2021/22 then recovered",
                             color_discrete_map={la:BOROUGH_COLOURS.get(la,WCC["blue"]) for la in ts_agg["la"].unique()})
            fig_ts.update_traces(line_width=2.5)
            fig_ts.update_xaxes(title="Academic year")
            fig_ts.update_yaxes(title="Avg Attainment 8 score")
            apply_wcc_style(fig_ts, "DfE KS4 Performance, Explore Education Statistics")
            st.plotly_chart(fig_ts, use_container_width=True)
            pptx_btn(fig_ts, "att8_trend", "Attainment 8 trend 2018/19–2024/25")

    st.markdown('<div class="source-box">DfE Key Stage 4 attainment by ethnicity, state-funded schools, Inner London. 2018/19–2024/25. Note: 2020/21 and 2021/22 figures should be interpreted with caution due to COVID-19 assessment disruption. Small ethnic group numbers lead to suppressed cells.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ETHNICITY & DEMOGRAPHICS
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Children's ethnicity & age structure — Westminster LSOAs (Census 2021)")
    col_e1, col_e2 = st.columns(2)

    with col_e1:
        eth_totals = {"Asian":df_rm033["Asian"].sum(),"Black":df_rm033["Black"].sum(),
                      "Mixed":df_rm033["Mixed"].sum(),"White":df_rm033["White"].sum(),
                      "Arab":df_rm033["Arab"].sum(),"Other":df_rm033["Other"].sum()}
        eth_df = pd.DataFrame(list(eth_totals.items()),columns=["Group","Count"])
        eth_df = eth_df.sort_values("Count",ascending=True)
        eth_df["Pct"] = eth_df["Count"]/eth_df["Count"].sum()*100
        fig_eb = px.bar(eth_df, x="Count", y="Group", orientation="h",
                        title="White (30%) and Asian (24%) are largest groups",
                        color="Group",
                        color_discrete_sequence=[WCC["blue"],WCC["cobalt"],WCC["amaranth"],
                                                 WCC["yellow"],WCC["green"],WCC["orange"]],
                        text="Pct")
        fig_eb.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig_eb.update_xaxes(title="Dependent children")
        fig_eb.update_yaxes(title="")
        apply_wcc_style(fig_eb, "Census 2021, ONS Nomis RM033")
        st.plotly_chart(fig_eb, use_container_width=True)
        pptx_btn(fig_eb, "wcc_ethnic", "Westminster children by ethnic group (Census 2021)")

    with col_e2:
        age_totals = {"0–4":df_rm006["Age_0_4"].sum(),"5–9":df_rm006["Age_5_9"].sum(),
                      "10–15":df_rm006["Age_10_15"].sum(),"16–18":df_rm006["Age_16_18"].sum()}
        age_df = pd.DataFrame(list(age_totals.items()),columns=["Age","Count"])
        age_df["Pct"] = age_df["Count"]/age_df["Count"].sum()*100
        fig_age = px.bar(age_df, x="Age", y="Pct",
                         title="10–15 is the largest age group (36%)",
                         color="Age",
                         color_discrete_sequence=[WCC["blue"],WCC["cobalt"],WCC["amaranth"],WCC["orange"]],
                         text="Pct")
        fig_age.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig_age.update_yaxes(rangemode="tozero", title="% of dependent children")
        fig_age.update_xaxes(title="Age group")
        apply_wcc_style(fig_age, "Census 2021, ONS Nomis RM006")
        st.plotly_chart(fig_age, use_container_width=True)
        pptx_btn(fig_age, "wcc_age", "Age structure of Westminster's children (Census 2021)")

    st.divider()
    st.subheader("Ethnic diversity vs number of children by LSOA")
    df_sc = df_rm033.copy()
    df_sc["diversity_idx"] = 1 - (
        (df_sc["White"]/df_sc["Total"])**2 + (df_sc["Asian"]/df_sc["Total"])**2 +
        (df_sc["Black"]/df_sc["Total"])**2 + (df_sc["Mixed"]/df_sc["Total"])**2 +
        (df_sc["Other"]/df_sc["Total"])**2)
    df_sc = df_sc.merge(df_rm006[["LSOA_CODE","Total_dep_children"]],on="LSOA_CODE",how="left")
    df_sc = df_sc[df_sc["Total"]>10]
    fig_sc = px.scatter(df_sc, x="Total", y="diversity_idx", hover_name="LSOA_NAME",
                        size="Total", color="diversity_idx", color_continuous_scale="Blues",
                        title="Larger LSOAs tend to have higher ethnic diversity",
                        labels={"Total":"Total dependent children","diversity_idx":"Diversity index"},
                        trendline="ols")
    apply_wcc_style(fig_sc, "Census 2021, ONS Nomis RM033 — diversity = 1 − Σ(share²)")
    st.plotly_chart(fig_sc, use_container_width=True)
    pptx_btn(fig_sc, "diversity_scatter", "Ethnic diversity vs children by Westminster LSOA")

    st.markdown('<div class="source-box">Census 2021, ONS Nomis. RM006: Age of youngest dependent child by household type. RM033: Ethnic group of dependent child by sex. Both at 2021 LSOA level.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — EGDI
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("Ethnic Group Deprivation Index (EGDI) — Westminster & CIPFA neighbours")
    st.caption("EGDI measures how unevenly deprivation is distributed across ethnic groups within a local authority. Source: Lloyd et al. (2023), gedi.ac.uk/egdi")

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
        fig_dec = px.line(df_dec.sort_values("Decile_n"),
                          x="Decile_n", y="Pct", color="LA_Name", markers=True,
                          color_discrete_map={la:BOROUGH_COLOURS.get(la,WCC["blue"]) for la in df_dec["LA_Name"].unique()},
                          title="Westminster: 22% of LSOAs in most deprived national decile",
                          labels={"Decile_n":"Income deprivation decile (1=most deprived)","Pct":"% of LSOAs"})
        fig_dec.update_traces(line_width=2.5)
        fig_dec.update_xaxes(tickvals=list(range(1,11)), title="Income deprivation decile")
        fig_dec.update_yaxes(rangemode="tozero", title="% of borough LSOAs in decile")
        apply_wcc_style(fig_dec, "EGDI (Lloyd et al. 2023) — gedi.ac.uk/egdi. Based on IMD 2019.")
        st.plotly_chart(fig_dec, use_container_width=True)
        pptx_btn(fig_dec, "egdi_decile", "EGDI deprivation decile distribution")

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

    # Radar chart
    st.divider()
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
            name=la, line=dict(color=BOROUGH_COLOURS.get(la,WCC["blue"]), width=2),
            fill="toself", fillcolor=BOROUGH_COLOURS.get(la,WCC["blue"]), opacity=0.12))
    fig_r.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0,30])),
        title="Westminster and Islington: most LSOAs in most deprived deciles",
        font_family="Arial", paper_bgcolor="white", height=440,
        legend=dict(orientation="h", y=-0.1))
    st.plotly_chart(fig_r, use_container_width=True)
    pptx_btn(fig_r, "egdi_radar", "EGDI radar chart — deprivation decile distribution")

    st.markdown('<div class="source-box">EGDI (Ethnic Group Deprivation Index), Lloyd et al. (2023). GEDI. <a href="https://gedi.ac.uk/egdi/">gedi.ac.uk/egdi</a>. Based on IMD 2019 income deprivation domain. "More ethnic inequality" means deprivation is significantly unequally distributed across ethnic groups within the borough — not that it is more deprived overall.</div>', unsafe_allow_html=True)
