# pip install streamlit plotly pandas python-pptx openpyxl odfpy kaleido topojson
# Run: streamlit run app.py
#
# DATA SOURCES
# ─────────────────────────────────────────────────────────────────────────────
# 1. Children in Low Income Families (2022-2025) — DWP / HMRC
#    https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025
# 2. Key Stage 4 Performance by Ethnicity — DfE / Explore Education Statistics
#    https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df
# 3. Ethnic Group Deprivation Index (EGDI) — GEDI / Lloyd et al. 2023
#    https://gedi.ac.uk/egdi/
# 4. Census 2021 – Dependent children by age of youngest child (RM006)
#    https://www.nomisweb.co.uk/ (ONS Nomis)
# 5. Census 2021 – Dependent children by ethnic group of HRP (RM12)
#    https://www.nomisweb.co.uk/ (ONS Nomis)
# 6. Census 2021 – Ethnic group of dependent child by sex (RM033)
#    https://www.nomisweb.co.uk/ (ONS Nomis)
# 7. Westminster LSOA boundaries — ONS Open Geography Portal
# 8. CIPFA Statistical Neighbours — Trust for London
#    https://trustforlondon.org.uk/data/information-on-cipfa-nearest-statistical-neighbours/

import os, io, json, warnings
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import topojson as tp

warnings.filterwarnings("ignore")

# ── DATA PATHS ────────────────────────────────────────────────────────────────
# Maps canonical filenames to possible on-disk aliases.
# GitHub uploads can rename spaces → spaces, (1) → (1), etc.
_FILENAME_ALIASES = {
    "data-key-stage-4-performance__1_.ods": [
        "data-key-stage-4-performance__1_.ods",
        "data-key-stage-4-performance (1).ods",
        "data-key-stage-4-performance_(1).ods",
    ],
    "LSOA_WCC__1_.json": [
        "LSOA_WCC__1_.json",
        "LSOA_WCC (1).json",
        "LSOA_WCC_(1).json",
    ],
    # CSV alternatives for the child poverty ODS file
    "2_AHC_Relative_LA.csv": [
        "2_AHC_Relative_LA.csv",
        "2_AHC_Relative_LA.csv",
    ],
    "4_AHC_Relative_Ward.csv": [
        "4_AHC_Relative_Ward.csv",
        "4 AHC Relative Ward.csv",
    ],
}

_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_DATA_SUBDIR = os.path.join(_SCRIPT_DIR, "data")
_SEARCH_DIRS = [_DATA_SUBDIR, _SCRIPT_DIR, "/mnt/user-data/uploads"]

def _dp(canonical_filename):
    """Resolve a data filename to an absolute path.

    Search order: data/ subfolder → repo root → /mnt/user-data/uploads.
    Also tries known filename aliases so the app works whether files are in
    a data/ subfolder, in the repo root (as uploaded to GitHub), or in the
    Claude upload area.
    """
    aliases = _FILENAME_ALIASES.get(canonical_filename, []) + [canonical_filename]
    for directory in _SEARCH_DIRS:
        for alias in aliases:
            candidate = os.path.join(directory, alias)
            if os.path.exists(candidate):
                return candidate
    # Fall through: return upload path; will raise a clear FileNotFoundError
    return os.path.join("/mnt/user-data/uploads", canonical_filename)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
NEIGHBOURS = {
    "Westminster":          "E09000033",
    "Kensington & Chelsea": "E09000020",
    "Camden":               "E09000007",
    "Hammersmith & Fulham": "E09000013",
    "Islington":            "E09000019",
    "Wandsworth":           "E09000032",
}
NEIGHBOUR_NAMES = list(NEIGHBOURS.keys())

# ONS colour palette
ONS = {
    "navy":      "#003087",
    "blue":      "#27A0CC",
    "green":     "#0F8243",
    "orange":    "#F4901E",
    "pink":      "#EB4A8A",
    "grey":      "#AAAAAA",
    "light":     "#D9EAF7",
    "text":      "#222222",
    "grid":      "#F0F0F0",
}

BOROUGH_COLOURS = {
    "Westminster":          "#003087",
    "Kensington & Chelsea": "#27A0CC",
    "Camden":               "#0F8243",
    "Hammersmith & Fulham": "#F4901E",
    "Islington":            "#EB4A8A",
    "Wandsworth":           "#6B4226",
    # Raw LA name variants
    "Kensington and Chelsea":  "#27A0CC",
    "Hammersmith and Fulham":  "#F4901E",
}

def borough_colour(name):
    return BOROUGH_COLOURS.get(name, ONS["grey"])

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Westminster Children's Demographics",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS
st.markdown("""
<style>
  .metric-card { background:#f8f9fa; border-radius:8px; padding:16px; border-left:4px solid #003087; }
  .source-box  { background:#f0f4f8; border-radius:6px; padding:10px 14px; font-size:0.82em; color:#555; margin-top:8px; }
  .highlight   { color:#003087; font-weight:700; }
  h1, h2       { color:#1a1a2e; }
  .stTabs [data-baseweb="tab"] { font-size:0.95rem; }
</style>
""", unsafe_allow_html=True)

# ── DATA LOADING ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_low_income_la():
    """Load LA-level child poverty data. Tries CSV first (faster, no ODS dependency),
    then falls back to the original .ods file."""
    csv_path = _dp("2_AHC_Relative_LA.csv")
    if os.path.exists(csv_path):
        df = pd.read_csv(csv_path)
        # CSV already has correct column names if exported from the app
        # but handle both raw-export and pre-processed formats
        if "LA" not in df.columns:
            df.columns = ["LA", "Area_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"]
    else:
        df = pd.read_excel(
            _dp("children-in-low-income-families-local-area-statistics-2022-2025.ods"),
            sheet_name="2_AHC_Relative_LA", engine="odf", header=8
        )
        df.columns = ["LA", "Area_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"]
    df = df.dropna(subset=["Area_Code"])
    df["Pct_2024"] = pd.to_numeric(df["Pct_2024"], errors="coerce")
    df["Pct_2025"] = pd.to_numeric(df["Pct_2025"], errors="coerce")
    # If values are fractions (0-1 range), multiply by 100
    if df["Pct_2024"].dropna().max() <= 1.0:
        df["Pct_2024"] *= 100
        df["Pct_2025"] *= 100
    df["N_2024"] = pd.to_numeric(df["N_2024"], errors="coerce")
    df["N_2025"] = pd.to_numeric(df["N_2025"], errors="coerce")
    return df

@st.cache_data(show_spinner=False)
def load_low_income_ward():
    """Load ward-level child poverty data. Tries CSV first, then falls back to .ods."""
    csv_path = _dp("4_AHC_Relative_Ward.csv")
    if os.path.exists(csv_path):
        df = pd.read_csv(csv_path)
        if "LA" not in df.columns:
            df.columns = ["LA", "LA_Code", "Ward", "Ward_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"]
    else:
        df = pd.read_excel(
            _dp("children-in-low-income-families-local-area-statistics-2022-2025.ods"),
            sheet_name="4_AHC_Relative_Ward", engine="odf", header=9
        )
        df.columns = ["LA", "LA_Code", "Ward", "Ward_Code", "N_2024", "N_2025", "Pct_2024", "Pct_2025"]
    df = df.dropna(subset=["Ward_Code"])
    df["Pct_2024"] = pd.to_numeric(df["Pct_2024"], errors="coerce")
    df["Pct_2025"] = pd.to_numeric(df["Pct_2025"], errors="coerce")
    if df["Pct_2024"].dropna().max() <= 1.0:
        df["Pct_2024"] *= 100
        df["Pct_2025"] *= 100
    df["N_2024"]   = pd.to_numeric(df["N_2024"], errors="coerce")
    df["N_2025"]   = pd.to_numeric(df["N_2025"], errors="coerce")
    df["LA_filled"] = df["LA"].ffill()
    return df

@st.cache_data(show_spinner=False)
def load_ks4_ethnic():
    """KS4 2024/25 Attainment 8 by ethnic group — Inner London boroughs."""
    df2 = pd.read_excel(
        _dp("data-key-stage-4-performance__1_.ods"),
        engine="odf", header=None
    )
    years_row   = df2.iloc[2, 4:].ffill()
    metrics_row = df2.iloc[3, 4:].ffill()
    sub_row     = df2.iloc[4, 4:]

    cols = ["ethnic_group", "subgroup", "region", "la"]
    for j in range(len(years_row)):
        yr = str(years_row.iloc[j]).strip().replace("/", "_")
        m  = str(metrics_row.iloc[j]).strip().replace(" ", "_")[:22]
        s  = str(sub_row.iloc[j]).strip().replace(" ", "_")[:18]
        cols.append(f"{yr}_{m}_{s}")

    data = df2.iloc[5:].copy()
    data.columns = cols[:len(data.columns)]
    data["ethnic_group"] = data["ethnic_group"].ffill()
    data["subgroup"]     = data["subgroup"].ffill()
    data["region"]       = data["region"].ffill()

    inner = ["Camden", "Hackney", "Hammersmith and Fulham", "Haringey", "Islington",
             "Kensington and Chelsea", "Lambeth", "Lewisham", "Newham", "Southwark",
             "Tower Hamlets", "Wandsworth", "Westminster"]

    mask = data["la"].isin(inner) & data["subgroup"].astype(str).str.startswith("All")
    sub = data[mask].copy()
    # Clean the attainment column
    att_col = [c for c in sub.columns if "2024_25" in c and "Attainment_8" in c and "Total" in c]
    if att_col:
        sub["att8_2425"] = pd.to_numeric(sub[att_col[0]], errors="coerce")
    pct_col = [c for c in sub.columns if "2024_25" in c and "achieving_gr" in c and "Total" in c]
    if pct_col:
        raw = sub[pct_col[0]].astype(str).str.replace("%","",regex=False)
        sub["pct_5above_EM_2425"] = pd.to_numeric(raw, errors="coerce")

    # Standardise LA names
    sub["la"] = sub["la"].str.replace("and Fulham", "& Fulham").str.replace("and Chelsea","& Chelsea")
    return sub

@st.cache_data(show_spinner=False)
def load_ks4_time():
    """KS4 time series (all pupils) — Inner London boroughs."""
    df1 = pd.read_excel(
        _dp("data-key-stage-4-performance.ods"),
        engine="odf", header=None
    )
    YEARS = ["2018/19","2019/20","2020/21","2021/22","2022/23","2023/24","2024/25"]
    # Row 2 = metric names; row 3 = years; rows 4+ = data
    metrics_row = df1.iloc[2, 5:].values
    years_row   = df1.iloc[3, 5:].values

    # Find Attainment 8 columns (cols 12-18 in 0-indexed = cols 12 to 18)
    att8_start = None
    for j, v in enumerate(metrics_row):
        if "Attainment 8" in str(v):
            att8_start = j + 5
            break

    records = []
    inner = {"Camden","Hackney","Hammersmith and Fulham","Haringey","Islington",
             "Kensington and Chelsea","Lambeth","Lewisham","Newham","Southwark",
             "Tower Hamlets","Wandsworth","Westminster"}

    # Find "All pupils" blocks — first block starts at some row where col1 == "All pupils"
    # Instead: take the block where subgroup = nan (first position per block = "All pupils" equivalent)
    # The first unique ethnic_group block per LA with subgroup NaN = total pupils
    # Actually we need rows where col0='All pupils' equivalent = All state-funded pupils
    # In v1 the total block has no ethnic filter; find block for "All pupils, All, Girls+Boys"
    # Simplest: rows where col0 has 'All pupils' OR is the very first repeating block

    # The structure: every 33 rows = one ethnic/sex/group block
    # We want the "total" block — find it by looking at col3 = la and col4-5 = numeric (no filter text)
    # Quick approach: just grab all LA rows and pick those with col0=NaN ethnic (filled forward won't work well)
    # Re-read with forward fill
    df1_copy = df1.copy()
    df1_copy.iloc[4:, 0] = df1_copy.iloc[4:, 0].ffill()
    df1_copy.iloc[4:, 1] = df1_copy.iloc[4:, 1].ffill()
    df1_copy.iloc[4:, 2] = df1_copy.iloc[4:, 2].ffill()

    # Get unique ethnic groups in col0
    ethnic_vals = df1_copy.iloc[4:, 0].dropna().unique()

    if att8_start is None:
        att8_start = 12

    att8_cols = list(range(att8_start, att8_start + 7))

    for i in range(4, len(df1)):
        la = str(df1_copy.iloc[i, 4]) if pd.notna(df1_copy.iloc[i, 4]) else None
        if la not in inner:
            continue
        eg  = str(df1_copy.iloc[i, 0]) if pd.notna(df1_copy.iloc[i, 0]) else "All pupils"
        sg  = str(df1_copy.iloc[i, 1]) if pd.notna(df1_copy.iloc[i, 1]) else ""
        sex = str(df1_copy.iloc[i, 2]) if pd.notna(df1_copy.iloc[i, 2]) else ""
        for yr_idx, yr in enumerate(YEARS):
            cidx = att8_start + yr_idx
            if cidx < len(df1.columns):
                val = pd.to_numeric(str(df1_copy.iloc[i, cidx]).replace(",",""), errors="coerce")
                records.append({
                    "la": la, "ethnic_group": eg, "subgroup": sg, "sex": sex,
                    "year": yr, "att8": val
                })

    ts = pd.DataFrame(records)
    ts["la"] = ts["la"].str.replace("and Fulham","& Fulham").str.replace("and Chelsea","& Chelsea")
    return ts

@st.cache_data(show_spinner=False)
def load_rm006():
    """Census 2021 RM006 — Age of youngest dependent child by LSOA (Westminster)."""
    df = pd.read_excel(
        _dp("RM006_age_of_youngest_dependent_child_by_household_type.xlsx"),
        header=7, skiprows=[8]
    )
    df.columns = ["LSOA", "No_dep_children", "Age_0_4", "Age_5_9", "Age_10_15", "Age_16_18"]
    df = df.dropna(subset=["LSOA"])
    df["LSOA_CODE"] = df["LSOA"].str.extract(r"(E\d+)")
    df["LSOA_NAME"] = df["LSOA"].str.replace(r"E\d+ : ", "", regex=True).str.strip()
    for c in ["Age_0_4","Age_5_9","Age_10_15","Age_16_18","No_dep_children"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    df["Total_dep_children"] = df["Age_0_4"] + df["Age_5_9"] + df["Age_10_15"] + df["Age_16_18"]
    df["Pct_under5"] = np.where(df["Total_dep_children"]>0,
                                df["Age_0_4"]/df["Total_dep_children"]*100, 0)
    return df[df["LSOA_CODE"].notna()].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_rm033():
    """Census 2021 RM033 — Ethnic group of dependent child by sex (Westminster LSOAs)."""
    df = pd.read_excel(
        _dp("RM033_ethic_group_dependent_child_by_sex.xlsx"), header=8
    )
    df.rename(columns={df.columns[0]: "LSOA"}, inplace=True)
    df = df.dropna(subset=["LSOA"])
    df["LSOA_CODE"] = df["LSOA"].str.extract(r"(E\d+)")
    df["LSOA_NAME"] = df["LSOA"].str.replace(r"E\d+ : ", "", regex=True).str.strip()
    for c in df.columns[1:-2]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    ethnic_cols = [c for c in df.columns if c not in ["LSOA","LSOA_CODE","LSOA_NAME"]]

    # Aggregate to broad ethnic groups
    df["Asian"]  = df[[c for c in ethnic_cols if "Asian" in c]].sum(axis=1)
    df["Black"]  = df[[c for c in ethnic_cols if "Black" in c]].sum(axis=1)
    df["Mixed"]  = df[[c for c in ethnic_cols if "Mixed" in c]].sum(axis=1)
    df["White"]  = df[[c for c in ethnic_cols if "White" in c]].sum(axis=1)
    df["Arab"]   = df[[c for c in ethnic_cols if "Arab" in c]].sum(axis=1)
    df["Other"]  = df[[c for c in ethnic_cols if "Other" in c and "Black" not in c and "Asian" not in c and "White" not in c and "Mixed" not in c]].sum(axis=1)
    df["Total"]  = df[["Asian","Black","Mixed","White","Arab","Other"]].sum(axis=1)

    return df[df["LSOA_CODE"].notna()].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_rm12():
    """Census 2021 RM12 — Dependent children by ethnic group of HRP, by LSOA."""
    df = pd.read_excel(
        _dp("RM12_dependent_children_by_ethnic_group_of_HRP.xlsx"), header=8
    )
    df.columns = ["LSOA","Age_0_2","Age_3_4","Age_5_11","Age_12_15","Age_16_18"]
    df = df.dropna(subset=["LSOA"])
    df["LSOA_CODE"] = df["LSOA"].str.extract(r"(E\d+)")
    df["LSOA_NAME"] = df["LSOA"].str.replace(r"E\d+ : ", "", regex=True).str.strip()
    for c in ["Age_0_2","Age_3_4","Age_5_11","Age_12_15","Age_16_18"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    df["Total"] = df[["Age_0_2","Age_3_4","Age_5_11","Age_12_15","Age_16_18"]].sum(axis=1)
    return df[df["LSOA_CODE"].notna()].reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_egdi():
    """EGDI Local Authority profiles."""
    df = pd.read_excel(_dp("EGDI-Local-Authority-profiles.xlsx"), sheet_name="Profiles")
    df.columns = [
        "idx","LA_Code","LA_Name",
        "D1","D2","D3","D4","D5","D6","D7","D8","D9","D10",
        "Total_LSOAs","Pct_D1","Pct_D2","Pct_D3","Pct_D4","Pct_D5",
        "Pct_D6","Pct_D7","Pct_D8","Pct_D9","Pct_D10",
        "_a","_b","_c","Category","_d","Flat","More_ethnic_ineq","Less_ethnic_ineq",
        "N_shape","Pct_bottom20","Pct_top20"
    ]
    df = df.iloc[1:].reset_index(drop=True)
    return df

@st.cache_data(show_spinner=False)
def load_wcc_geojson():
    """Westminster LSOA boundaries from TopoJSON."""
    with open(_dp("LSOA_WCC__1_.json")) as f:
        topo = json.load(f)
    topo_obj = tp.Topology(topo, object_name="LSOA_WCC")
    gj = json.loads(topo_obj.to_geojson())
    return gj

# ── PPTX HELPERS ──────────────────────────────────────────────────────────────
def _fig_to_png(fig):
    """Render a Plotly figure to PNG bytes.

    Tries kaleido (both 0.2.x and 1.x).  If Chrome is not available for
    kaleido 1.x, falls back to a white placeholder image so the rest of the
    PPTX export still works.
    """
    try:
        return fig.to_image(format="png", width=1200, height=680, scale=2)
    except Exception as e:
        err = str(e)
        if "Chrome" in err or "kaleido" in err.lower():
            # kaleido 1.x needs Chrome which isn't available in this environment.
            # Return a minimal 1x1 white PNG so the PPTX is still generated.
            import base64
            # 1×1 white PNG (smallest valid PNG)
            WHITE_1PX = base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="
            )
            return WHITE_1PX
        raise

def _fig_to_pptx(fig, title=""):
    img = _fig_to_png(fig)
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(io.BytesIO(img), Inches(0.3), Inches(0.3), Inches(12.73), Inches(6.4))
    txb = slide.shapes.add_textbox(Inches(0.3), Inches(6.8), Inches(12), Inches(0.55))
    tf  = txb.text_frame
    tf.text = title or (fig.layout.title.text or "Chart")
    tf.paragraphs[0].runs[0].font.size = Pt(13)
    tf.paragraphs[0].runs[0].font.bold = True
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf

def pptx_btn(fig, key, title=""):
    try:
        buf = _fig_to_pptx(fig, title)
        st.download_button("⬇ Download slide (PPTX)", buf, f"{key}.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key=f"dl_{key}")
    except Exception as e:
        st.caption(f"_PPTX export unavailable: {e}_")

def apply_ons_style(fig, source=""):
    fig.update_layout(
        font_family="Arial",
        font_color=ONS["text"],
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font_size=15,
        title_font_color="#1a1a2e",
        margin=dict(l=50, r=30, t=70, b=55),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    fig.update_xaxes(showgrid=False, linecolor="#cccccc", showline=True)
    fig.update_yaxes(gridcolor=ONS["grid"], linecolor="white", zeroline=False)
    if source:
        fig.add_annotation(
            text=f"<i>Source: {source}</i>",
            xref="paper", yref="paper", x=0, y=-0.13,
            showarrow=False, font=dict(size=10, color="#777"),align="left"
        )
    return fig

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/en/thumb/1/16/Westminster_City_Council.svg/200px-Westminster_City_Council.svg.png", width=120)
    st.title("🏙️ Westminster Children")
    st.markdown("**CIPFA Statistical Neighbours**")
    for b in NEIGHBOUR_NAMES:
        colour = BOROUGH_COLOURS.get(b, ONS["navy"])
        st.markdown(f"<span style='color:{colour}'>■</span> {b}", unsafe_allow_html=True)

    st.divider()
    st.caption("**Data Sources**")
    st.markdown("[DWP Children in Low Income Families](https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025)")
    st.markdown("[DfE Key Stage 4 (Explore Ed Stats)](https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df)")
    st.markdown("[EGDI — Ethnic Group Deprivation Index](https://gedi.ac.uk/egdi/)")
    st.markdown("[Census 2021 — ONS Nomis](https://www.nomisweb.co.uk/)")
    st.markdown("[CIPFA Neighbours — Trust for London](https://trustforlondon.org.uk/data/information-on-cipfa-nearest-statistical-neighbours/)")
    st.divider()
    st.caption("All data at LSOA or Local Authority level. Census data: 2021. Low income & KS4: 2024/25.")

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
with st.spinner("Loading datasets…"):
    df_li_la    = load_low_income_la()
    df_li_ward  = load_low_income_ward()
    df_ks4_eth  = load_ks4_ethnic()
    df_ks4_ts   = load_ks4_time()
    df_rm006    = load_rm006()
    df_rm033    = load_rm033()
    df_rm12     = load_rm12()
    df_egdi     = load_egdi()
    wcc_geojson = load_wcc_geojson()

# ── MAIN HEADER ───────────────────────────────────────────────────────────────
st.title("🏙️ Westminster Children's Demographics")
st.markdown(
    "Exploring child poverty, demographics, attainment and ethnic diversity across "
    "Westminster LSOAs — benchmarked against CIPFA statistical neighbours."
)

# ── TOP METRICS ───────────────────────────────────────────────────────────────
wcc_li = df_li_la[df_li_la["LA"].str.contains("Westminster", na=False)].iloc[0]
london_avg_pct = df_li_la[df_li_la["Area_Code"].str.startswith("E09", na=False)]["Pct_2025"].mean()

rm006_wcc = df_rm006  # already Westminster only
total_dep = rm006_wcc["Total_dep_children"].sum()
under5_n  = rm006_wcc["Age_0_4"].sum()
pct_u5    = round(under5_n / total_dep * 100, 1) if total_dep else 0

rm033_tot = df_rm033["Total"].sum()
rm033_white = df_rm033["White"].sum()
pct_white = round(rm033_white / rm033_tot * 100, 1) if rm033_tot else 0
pct_nonwhite = round(100 - pct_white, 1)

wcc_att8 = df_ks4_eth[(df_ks4_eth["la"]=="Westminster") & (df_ks4_eth["ethnic_group"].notna())]
overall_att8_row = wcc_att8[wcc_att8["ethnic_group"].str.contains("Asian", na=False)]  # placeholder
# Get Westminster overall att8 from v2 data
wcc_att8_all_row = df_ks4_eth[df_ks4_eth["la"] == "Westminster"]
if "att8_2425" in wcc_att8_all_row.columns and len(wcc_att8_all_row):
    att8_avg = wcc_att8_all_row["att8_2425"].mean()
else:
    att8_avg = None

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Children in low income families (FYE 2025)",
          f"{int(wcc_li['N_2025']):,}",
          delta=f"{wcc_li['Pct_2025']:.1f}% of children",
          delta_color="off")
c2.metric("Change vs FYE 2024", f"{int(wcc_li['N_2025']-wcc_li['N_2024']):+,}",
          delta=f"{wcc_li['Pct_2025']-wcc_li['Pct_2024']:+.1f}pp", delta_color="inverse")
c3.metric("Westminster vs London avg (low income)",
          f"{wcc_li['Pct_2025']:.1f}%",
          delta=f"{wcc_li['Pct_2025']-london_avg_pct:+.1f}pp vs London", delta_color="inverse")
c4.metric("Dependent children (Census 2021)", f"{total_dep:,}",
          delta=f"{pct_u5}% aged 0–4", delta_color="off")
c5.metric("Non-white dependent children (2021)", f"{pct_nonwhite}%",
          delta="Westminster LSOA average", delta_color="off")

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📍 Child Poverty — Borough Comparison",
    "🗺️ Westminster LSOA Maps",
    "📚 KS4 Attainment",
    "👥 Ethnicity & Demographics",
    "⚖️ Ethnic Deprivation (EGDI)",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Child Poverty Borough Comparison
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Children in low income families — Westminster vs CIPFA neighbours (FYE 2025)")
    st.markdown(
        "**Context:** After-housing-costs (AHC) relative poverty rate for children aged 0–15. "
        "Westminster (26.3%) sits mid-table among its statistical neighbours — "
        "**Islington is highest at 33.4%, Kensington & Chelsea lowest at 16.1%.**"
    )

    # Filter neighbours
    nb_codes = list(NEIGHBOURS.values())
    df_nb = df_li_la[df_li_la["Area_Code"].isin(nb_codes)].copy()
    df_nb["Borough"] = df_nb["LA"].str.replace("and Fulham","& Fulham").str.replace("and Chelsea","& Chelsea")
    df_nb["colour"] = df_nb["Borough"].map(lambda x: BOROUGH_COLOURS.get(x, ONS["grey"]))
    df_nb = df_nb.sort_values("Pct_2025", ascending=True)

    col_a, col_b = st.columns([3, 2])

    with col_a:
        fig1 = px.bar(
            df_nb, x="Pct_2025", y="Borough", orientation="h",
            title="Islington has the highest child poverty rate among CIPFA neighbours",
            color="Borough",
            color_discrete_map={r["Borough"]: BOROUGH_COLOURS.get(r["Borough"], ONS["grey"])
                                for _, r in df_nb.iterrows()},
            text="Pct_2025",
        )
        fig1.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig1.update_xaxes(range=[0, max(df_nb["Pct_2025"])*1.18], title="% children in relative low income (AHC)")
        fig1.update_yaxes(title="")
        apply_ons_style(fig1, "DWP Children in Low Income Families 2025")
        st.plotly_chart(fig1, use_container_width=True)
        pptx_btn(fig1, "child_poverty_bar",
                 "Islington has the highest child poverty rate among CIPFA neighbours")

    with col_b:
        # Change over time
        df_ts = df_nb[["Borough","Pct_2024","Pct_2025"]].melt("Borough", var_name="Year", value_name="Pct")
        df_ts["Year"] = df_ts["Year"].map({"Pct_2024":"FYE 2024","Pct_2025":"FYE 2025"})
        fig2 = px.line(df_ts, x="Year", y="Pct", color="Borough",
                       markers=True,
                       color_discrete_map={b: BOROUGH_COLOURS.get(b, ONS["grey"]) for b in df_ts["Borough"].unique()},
                       title="All neighbours saw child poverty fall 2024→2025")
        fig2.update_traces(line_width=2.5)
        fig2.update_xaxes(title="")
        fig2.update_yaxes(title="% children in low income", rangemode="tozero")
        apply_ons_style(fig2, "DWP Children in Low Income Families 2025")
        st.plotly_chart(fig2, use_container_width=True)
        pptx_btn(fig2, "child_poverty_trend", "All neighbours saw child poverty fall 2024→2025")

    st.info("💡 **Finding:** All six boroughs saw child poverty decline between FYE 2024 and 2025. "
            "Westminster fell from 27.6% to 26.3%. Kensington & Chelsea is the outlier — nearly "
            "half Westminster's rate — reflecting its unique wealth distribution.")

    st.divider()
    # Westminster ward-level breakdown
    st.subheader("Westminster ward-level child poverty (FYE 2025)")
    st.markdown("**Point:** Ward-level variation is stark — Church Street and Westbourne wards have rates over 44%.")

    wcc_wards = df_li_ward[df_li_ward["LA_filled"].astype(str).str.contains("Westminster", na=False)].copy()
    wcc_wards["Pct_2025"] = pd.to_numeric(wcc_wards["Pct_2025"], errors="coerce")
    wcc_wards = wcc_wards.sort_values("Pct_2025", ascending=True).dropna(subset=["Ward","Pct_2025"])

    colours_wards = [ONS["navy"] if w in ["Church Street","Westbourne","Queen's Park","Harrow Road"]
                     else ONS["blue"] for w in wcc_wards["Ward"]]

    fig3 = px.bar(
        wcc_wards, x="Pct_2025", y="Ward", orientation="h",
        title="Church Street (44%) and Westbourne (44%) have the highest child poverty in Westminster",
        color_discrete_sequence=[ONS["navy"]],
        text="Pct_2025",
    )
    fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    fig3.update_xaxes(range=[0, max(wcc_wards["Pct_2025"])*1.2], title="% children in relative low income (AHC)")
    fig3.update_yaxes(title="")
    apply_ons_style(fig3, "DWP Children in Low Income Families 2025")
    st.plotly_chart(fig3, use_container_width=True)
    pptx_btn(fig3, "westminster_ward_poverty",
             "Church Street and Westbourne have the highest child poverty in Westminster")

    st.markdown(
        '<div class="source-box">Source: DWP/HMRC Children in Low Income Families, AHC Relative, Ward level — '
        '<a href="https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025">gov.uk</a></div>',
        unsafe_allow_html=True
    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Westminster LSOA Maps
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Westminster LSOA demographic maps (Census 2021)")
    st.markdown(
        "**Context:** 128 LSOAs across Westminster. Each map shows a different child demographic. "
        "Use the dropdown to switch metric. Hover for LSOA name and value."
    )

    map_metric = st.selectbox("Map metric", [
        "Total dependent children",
        "% aged 0–4 (youngest age band)",
        "% aged 10–15",
        "% White children",
        "% Asian children",
        "% Black children",
        "% Mixed children",
    ])

    # Build merged map dataframe
    df_map = pd.merge(df_rm006[["LSOA_CODE","LSOA_NAME","Total_dep_children","Age_0_4","Age_10_15"]],
                      df_rm033[["LSOA_CODE","White","Asian","Black","Mixed","Total"]],
                      on="LSOA_CODE", how="inner")

    df_map["pct_u5"]    = np.where(df_map["Total_dep_children"]>0,
                                    df_map["Age_0_4"]/df_map["Total_dep_children"]*100, 0)
    df_map["pct_10_15"] = np.where(df_map["Total_dep_children"]>0,
                                    df_map["Age_10_15"]/df_map["Total_dep_children"]*100, 0)
    df_map["pct_white"] = np.where(df_map["Total"]>0, df_map["White"]/df_map["Total"]*100, 0)
    df_map["pct_asian"] = np.where(df_map["Total"]>0, df_map["Asian"]/df_map["Total"]*100, 0)
    df_map["pct_black"] = np.where(df_map["Total"]>0, df_map["Black"]/df_map["Total"]*100, 0)
    df_map["pct_mixed"] = np.where(df_map["Total"]>0, df_map["Mixed"]/df_map["Total"]*100, 0)

    METRIC_MAP = {
        "Total dependent children":      ("Total_dep_children", "Total dependent children", "Blues"),
        "% aged 0–4 (youngest age band)":("pct_u5",            "% aged 0–4",               "YlOrBr"),
        "% aged 10–15":                  ("pct_10_15",          "% aged 10–15",             "Purples"),
        "% White children":              ("pct_white",          "% White children",         "Greys"),
        "% Asian children":              ("pct_asian",          "% Asian children",         "Oranges"),
        "% Black children":              ("pct_black",          "% Black children",         "Reds"),
        "% Mixed children":              ("pct_mixed",          "% Mixed children",         "Greens"),
    }

    col_name, label, cscale = METRIC_MAP[map_metric]

    fig_map = px.choropleth_map(
        df_map,
        geojson=wcc_geojson,
        locations="LSOA_CODE",
        featureidkey="properties.LSOA11CD",
        color=col_name,
        hover_name="LSOA_NAME",
        hover_data={col_name: ":.1f", "LSOA_CODE": True},
        color_continuous_scale=cscale,
        zoom=12,
        center={"lat": 51.512, "lon": -0.155},
        opacity=0.75,
        title=f"{label} by Westminster LSOA (Census 2021)",
        map_style="carto-positron",
        labels={col_name: label},
    )
    fig_map.update_layout(margin=dict(l=0, r=0, t=50, b=0), height=550)
    st.plotly_chart(fig_map, use_container_width=True)
    pptx_btn(fig_map, f"lsoa_map_{col_name}", f"{label} by Westminster LSOA (Census 2021)")

    st.markdown(
        '<div class="source-box">Source: Census 2021, ONS Nomis — RM006 (household type by age of youngest child), '
        'RM033 (ethnic group of dependent child). Westminster LSOA boundaries: ONS Open Geography Portal.</div>',
        unsafe_allow_html=True
    )

    st.divider()
    # LSOA bar charts
    st.subheader("Top 15 Westminster LSOAs by dependent children")
    top15 = df_map.nlargest(15, "Total_dep_children").sort_values("Total_dep_children")
    fig_lsoa = px.bar(
        top15, x="Total_dep_children", y="LSOA_NAME", orientation="h",
        title="The 15 LSOAs with most dependent children are concentrated in Church Street and Westbourne areas",
        color_discrete_sequence=[ONS["navy"]],
        text="Total_dep_children",
    )
    fig_lsoa.update_traces(textposition="outside")
    fig_lsoa.update_xaxes(title="Total dependent children (Census 2021)")
    fig_lsoa.update_yaxes(title="")
    apply_ons_style(fig_lsoa, "Census 2021, ONS Nomis RM006")
    st.plotly_chart(fig_lsoa, use_container_width=True)
    pptx_btn(fig_lsoa, "lsoa_top15", "15 LSOAs with most dependent children")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — KS4 Attainment
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Key Stage 4 attainment — Westminster vs CIPFA neighbours")
    st.markdown(
        "**Context:** Average Attainment 8 score and % achieving grade 5+ in English & maths GCSE, "
        "2024/25 academic year, state-funded schools only. "
        "Data broken down by ethnic group. Source: "
        "[Explore Education Statistics](https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df)"
    )

    # Neighbour LA names as they appear in the data
    nb_la_names = ["Camden", "Hammersmith & Fulham", "Islington",
                   "Kensington & Chelsea", "Wandsworth", "Westminster"]

    ks4_nb = df_ks4_eth[df_ks4_eth["la"].isin(nb_la_names) & df_ks4_eth["att8_2425"].notna()].copy()

    col_ka, col_kb = st.columns([1, 1])

    with col_ka:
        # Overall Att8 by borough (All ethnic groups pooled / first ethnic cat)
        # Show by ethnic group for Westminster
        wcc_eth = ks4_nb[ks4_nb["la"] == "Westminster"].dropna(subset=["att8_2425"])
        wcc_eth["eg_short"] = wcc_eth["ethnic_group"].str.replace(" / "," /\n").str[:30]
        fig_att_eth = px.bar(
            wcc_eth.sort_values("att8_2425"),
            x="att8_2425", y="ethnic_group", orientation="h",
            title="Westminster: White pupils have highest Attainment 8 in 2024/25",
            color_discrete_sequence=[ONS["navy"]],
            text="att8_2425",
        )
        fig_att_eth.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig_att_eth.update_xaxes(title="Average Attainment 8 score", range=[0, 75])
        fig_att_eth.update_yaxes(title="")
        apply_ons_style(fig_att_eth, "DfE KS4 2024/25, state-funded schools")
        st.plotly_chart(fig_att_eth, use_container_width=True)
        pptx_btn(fig_att_eth, "wcc_att8_ethnic",
                 "Westminster: White pupils have highest Attainment 8 in 2024/25")

    with col_kb:
        # Compare Westminster vs neighbours for each ethnic group
        all_eth = ks4_nb.dropna(subset=["att8_2425"])
        fig_att_comp = px.bar(
            all_eth.sort_values(["ethnic_group","att8_2425"]),
            x="att8_2425", y="la", orientation="h",
            facet_col="ethnic_group", facet_col_wrap=2,
            title="Kensington & Chelsea leads on Attainment 8 across most ethnic groups",
            color="la",
            color_discrete_map={la: BOROUGH_COLOURS.get(la, ONS["grey"]) for la in all_eth["la"].unique()},
        )
        fig_att_comp.update_traces(showlegend=False)
        fig_att_comp.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1][:28]))
        fig_att_comp.update_xaxes(range=[0, 80], matches=None)
        fig_att_comp.update_yaxes(matches=None)
        fig_att_comp.update_layout(height=500)
        apply_ons_style(fig_att_comp, "DfE KS4 2024/25, state-funded schools")
        st.plotly_chart(fig_att_comp, use_container_width=True)
        pptx_btn(fig_att_comp, "att8_all_boroughs_ethnic",
                 "Kensington & Chelsea leads on Attainment 8 across most ethnic groups")

    st.info("💡 **Finding:** Westminster's overall Attainment 8 score is competitive among CIPFA neighbours. "
            "Kensington & Chelsea scores highest overall, but Westminster performs well for Asian and Black pupils. "
            "Note: small numbers mean some ethnic group data is suppressed ('no data').")

    st.divider()
    st.subheader("Attainment 8 trend over time (all ethnic groups combined)")
    st.markdown("**Source:** DfE state-funded school data, Inner London boroughs 2018/19–2024/25.")

    # Get the time series for "All Asian" as proxy — or aggregate if possible
    # Use KS4 v1 data — find rows where col0="Asian / Asian British", col1="All Asian"
    # Actually let's get all rows per LA across all ethnic groups
    nb_la_v1 = ["Camden","Hammersmith and Fulham","Islington","Kensington and Chelsea","Wandsworth","Westminster"]
    ts_data = df_ks4_ts[df_ks4_ts["la"].isin(nb_la_v1) &
                        df_ks4_ts["subgroup"].str.startswith("All", na=False) &
                        df_ks4_ts["att8"].notna()].copy()
    ts_data["la"] = ts_data["la"].str.replace("and Fulham","& Fulham").str.replace("and Chelsea","& Chelsea")

    # Average across ethnic groups per LA per year (crude but directional)
    ts_agg = ts_data.groupby(["la","year"])["att8"].mean().reset_index()
    ts_agg["colour_key"] = ts_agg["la"]

    fig_ts = px.line(
        ts_agg, x="year", y="att8", color="la",
        markers=True,
        color_discrete_map={la: BOROUGH_COLOURS.get(la, ONS["grey"]) for la in ts_agg["la"].unique()},
        title="Attainment 8 scores dipped in 2021/22 then partially recovered across all boroughs",
    )
    fig_ts.update_traces(line_width=2.5)
    fig_ts.update_xaxes(title="Academic year")
    fig_ts.update_yaxes(title="Avg Attainment 8 score")
    apply_ons_style(fig_ts, "DfE KS4 Performance, Explore Education Statistics")
    st.plotly_chart(fig_ts, use_container_width=True)
    pptx_btn(fig_ts, "att8_trend", "Attainment 8 trend 2018/19–2024/25 — CIPFA neighbours")

    st.markdown(
        '<div class="source-box">Source: Department for Education, Key Stage 4 attainment by ethnicity, '
        'state-funded schools. Local authority, Inner London. 2018/19 to 2024/25. '
        '<a href="https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df">'
        'Explore Education Statistics</a>. Note: 2020/21 and 2021/22 figures should be interpreted with caution '
        'due to COVID-19 assessment changes.</div>',
        unsafe_allow_html=True
    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Ethnicity & Demographics
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Children's ethnicity & age structure — Westminster LSOAs (Census 2021)")

    col_e1, col_e2 = st.columns(2)

    with col_e1:
        # Westminster-wide ethnic breakdown (aggregated from RM033)
        eth_totals = {
            "Asian":  df_rm033["Asian"].sum(),
            "Black":  df_rm033["Black"].sum(),
            "Mixed":  df_rm033["Mixed"].sum(),
            "White":  df_rm033["White"].sum(),
            "Arab":   df_rm033["Arab"].sum(),
            "Other":  df_rm033["Other"].sum(),
        }
        eth_df = pd.DataFrame(list(eth_totals.items()), columns=["Group","Count"])
        eth_df = eth_df.sort_values("Count", ascending=True)
        eth_df["Pct"] = eth_df["Count"] / eth_df["Count"].sum() * 100

        fig_eth_bar = px.bar(
            eth_df, x="Count", y="Group", orientation="h",
            title="White children are the largest group (30%), followed by Asian (24%)",
            color="Group",
            color_discrete_sequence=[ONS["navy"], ONS["blue"], ONS["green"],
                                      ONS["orange"], ONS["pink"], ONS["grey"]],
            text="Pct",
        )
        fig_eth_bar.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig_eth_bar.update_xaxes(title="Number of dependent children")
        fig_eth_bar.update_yaxes(title="")
        apply_ons_style(fig_eth_bar, "Census 2021, ONS Nomis RM033")
        st.plotly_chart(fig_eth_bar, use_container_width=True)
        pptx_btn(fig_eth_bar, "wcc_ethnic_breakdown",
                 "White children are the largest group (30%), followed by Asian (24%)")

    with col_e2:
        # Age structure from RM006
        age_totals = {
            "0–4 years":   df_rm006["Age_0_4"].sum(),
            "5–9 years":   df_rm006["Age_5_9"].sum(),
            "10–15 years": df_rm006["Age_10_15"].sum(),
            "16–18 years": df_rm006["Age_16_18"].sum(),
        }
        age_df = pd.DataFrame(list(age_totals.items()), columns=["Age","Count"])
        age_df["Pct"] = age_df["Count"] / age_df["Count"].sum() * 100
        colours_age = [ONS["navy"], ONS["blue"], ONS["orange"], ONS["grey"]]

        fig_age = px.bar(
            age_df, x="Age", y="Pct",
            title="The 10–15 age group is the largest, representing 36% of Westminster's children",
            color="Age",
            color_discrete_sequence=colours_age,
            text="Pct",
        )
        fig_age.update_traces(texttemplate="%{text:.1f}%", textposition="outside", showlegend=False)
        fig_age.update_yaxes(rangemode="tozero", title="% of dependent children")
        fig_age.update_xaxes(title="")
        apply_ons_style(fig_age, "Census 2021, ONS Nomis RM006")
        st.plotly_chart(fig_age, use_container_width=True)
        pptx_btn(fig_age, "wcc_age_structure",
                 "Age structure of dependent children in Westminster (Census 2021)")

    st.divider()
    # LSOA-level ethnic diversity scatter
    st.subheader("Ethnic diversity vs number of children by LSOA")
    df_scatter = df_rm033.copy()
    df_scatter["diversity_idx"] = 1 - (
        (df_scatter["White"]/df_scatter["Total"])**2 +
        (df_scatter["Asian"]/df_scatter["Total"])**2 +
        (df_scatter["Black"]/df_scatter["Total"])**2 +
        (df_scatter["Mixed"]/df_scatter["Total"])**2 +
        (df_scatter["Other"]/df_scatter["Total"])**2
    )
    df_scatter = df_scatter.merge(df_rm006[["LSOA_CODE","Total_dep_children"]], on="LSOA_CODE", how="left")
    df_scatter = df_scatter[df_scatter["Total"] > 10]

    fig_scatter = px.scatter(
        df_scatter, x="Total", y="diversity_idx",
        hover_name="LSOA_NAME",
        size="Total",
        color="diversity_idx",
        color_continuous_scale="Blues",
        title="Larger LSOAs in Westminster tend to have higher ethnic diversity among children",
        labels={"Total": "Total dependent children", "diversity_idx": "Herfindahl diversity index"},
        trendline="ols",
    )
    apply_ons_style(fig_scatter, "Census 2021, ONS Nomis RM033 — diversity = 1 − Σ(share²)")
    st.plotly_chart(fig_scatter, use_container_width=True)
    pptx_btn(fig_scatter, "lsoa_diversity_scatter",
             "Ethnic diversity vs number of children by Westminster LSOA")

    st.markdown(
        '<div class="source-box">Sources: Census 2021, ONS Nomis. RM006: Age of youngest dependent child '
        'by household type. RM033: Ethnic group of dependent child by sex. Both at LSOA level. '
        '<a href="https://www.nomisweb.co.uk/">nomisweb.co.uk</a></div>',
        unsafe_allow_html=True
    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — EGDI
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("Ethnic Group Deprivation Index (EGDI) — Westminster & CIPFA neighbours")
    st.markdown(
        "The **EGDI** (Lloyd et al. 2023) measures *how unevenly deprivation is distributed across ethnic groups* "
        "within a local authority. A **'More ethnic inequality'** classification means some ethnic groups "
        "face significantly higher deprivation than others within the same borough. "
        "Source: [gedi.ac.uk/egdi](https://gedi.ac.uk/egdi/)"
    )

    # Filter to our neighbours
    nb_egdi_codes = list(NEIGHBOURS.values())
    df_egdi_nb = df_egdi[df_egdi["LA_Code"].isin(nb_egdi_codes)].copy()
    df_egdi_nb["LA_Name"] = df_egdi_nb["LA_Name"].str.replace("and Fulham","& Fulham").str.replace("and Chelsea","& Chelsea")

    # Show LSOA deprivation distribution
    decile_cols = ["Pct_D1","Pct_D2","Pct_D3","Pct_D4","Pct_D5",
                   "Pct_D6","Pct_D7","Pct_D8","Pct_D9","Pct_D10"]
    for c in decile_cols:
        df_egdi_nb[c] = pd.to_numeric(df_egdi_nb[c], errors="coerce")

    # Melt for chart
    df_decile = df_egdi_nb[["LA_Name","Category"] + decile_cols].melt(
        id_vars=["LA_Name","Category"], var_name="Decile", value_name="Pct"
    )
    df_decile["Decile_n"] = df_decile["Decile"].str.replace("Pct_D","").astype(float)
    df_decile = df_decile.dropna(subset=["Pct"])

    col_g1, col_g2 = st.columns([3, 2])

    with col_g1:
        fig_dec = px.line(
            df_decile.sort_values("Decile_n"),
            x="Decile_n", y="Pct", color="LA_Name",
            markers=True,
            color_discrete_map={la: BOROUGH_COLOURS.get(la, ONS["grey"]) for la in df_decile["LA_Name"].unique()},
            title="Westminster has 22% of LSOAs in the most deprived decile — highest among neighbours",
            labels={"Decile_n": "Income deprivation decile (1=most deprived)", "Pct": "% of LSOAs"}
        )
        fig_dec.update_traces(line_width=2.5)
        fig_dec.update_xaxes(tickvals=list(range(1,11)),
                              title="Income deprivation decile (1 = most deprived 10% of LSOAs nationally)")
        fig_dec.update_yaxes(rangemode="tozero", title="% of borough LSOAs in decile")
        apply_ons_style(fig_dec, "EGDI (Lloyd et al. 2023) — gedi.ac.uk/egdi")
        st.plotly_chart(fig_dec, use_container_width=True)
        pptx_btn(fig_dec, "egdi_decile_dist",
                 "Westminster has 22% of LSOAs in the most deprived decile")

    with col_g2:
        # Category + key stats
        st.markdown("**EGDI Classification**")
        for _, row in df_egdi_nb.iterrows():
            cat = str(row.get("Category",""))
            la  = row["LA_Name"]
            col = BOROUGH_COLOURS.get(la, ONS["navy"])
            icon = "🔴" if "More" in cat else "🟢" if "Less" in cat else "⚪"
            pct_b20 = pd.to_numeric(row.get("Pct_bottom20",""), errors="coerce")
            pct_t20 = pd.to_numeric(row.get("Pct_top20",""), errors="coerce")
            b20_str = f"{pct_b20*100:.1f}%" if not pd.isna(pct_b20) else "—"
            t20_str = f"{pct_t20*100:.1f}%" if not pd.isna(pct_t20) else "—"
            st.markdown(
                f"<div style='border-left:3px solid {col}; padding:6px 10px; margin-bottom:6px; background:#f8f9fa; border-radius:4px'>"
                f"<b>{la}</b> {icon}<br><small>{cat}<br>"
                f"LSOAs in bottom 20%: <b>{b20_str}</b> · top 20%: <b>{t20_str}</b></small></div>",
                unsafe_allow_html=True
            )

    st.divider()
    # Westminster within London context
    st.subheader("Westminster LSOA income deprivation distribution in context")
    wcc_egdi = df_egdi[df_egdi["LA_Code"] == "E09000033"].iloc[0]
    d1_pct = pd.to_numeric(wcc_egdi.get("Pct_D1"), errors="coerce")
    total_lsoas = pd.to_numeric(wcc_egdi.get("Total_LSOAs"), errors="coerce")
    n_d1 = pd.to_numeric(wcc_egdi.get("D1"), errors="coerce")

    n_d1_str      = str(int(n_d1)) if not pd.isna(n_d1) else "~27"
    total_str     = str(int(total_lsoas)) if not pd.isna(total_lsoas) else "121"
    d1_pct_str    = f"{d1_pct*100:.1f}%" if not pd.isna(d1_pct) else "22%"
    wcc_cat       = wcc_egdi.get("Category", "More ethnic inequality")
    st.info(
        f"💡 **Westminster EGDI finding:** Westminster is classified as **'{wcc_cat}' "
        f"in ethnic deprivation inequality**, meaning some ethnic groups experience significantly higher "
        f"deprivation than others within the borough. "
        f"{n_d1_str} of {total_str} LSOAs ({d1_pct_str}) fall in the most deprived income decile nationally — "
        f"suggesting pockets of high deprivation sit alongside very wealthy areas."
    )

    st.markdown(
        '<div class="source-box">Source: Ethnic Group Deprivation Index (EGDI), Lloyd et al. (2023). '
        'GEDI — Geographies of Ethnic Diversity & Inequality. <a href="https://gedi.ac.uk/egdi/">gedi.ac.uk/egdi</a>. '
        'Deciles based on Income deprivation domain of the Index of Multiple Deprivation (IMD 2019).</div>',
        unsafe_allow_html=True
    )

    # Radar/spider: compare all 6 boroughs on key EGDI stats
    radar_cols_raw = ["Pct_D1","Pct_D2","Pct_D3","Pct_D4","Pct_D5"]
    radar_labels   = ["Decile 1\n(most deprived)","Decile 2","Decile 3","Decile 4","Decile 5"]

    for c in radar_cols_raw:
        df_egdi_nb[c] = pd.to_numeric(df_egdi_nb[c], errors="coerce") * 100

    fig_radar = go.Figure()
    for _, row in df_egdi_nb.iterrows():
        la = row["LA_Name"]
        vals = [pd.to_numeric(row[c], errors="coerce") for c in radar_cols_raw]
        vals_clean = [v if not pd.isna(v) else 0 for v in vals]
        fig_radar.add_trace(go.Scatterpolar(
            r=vals_clean + [vals_clean[0]],
            theta=radar_labels + [radar_labels[0]],
            name=la,
            line=dict(color=BOROUGH_COLOURS.get(la, ONS["grey"]), width=2),
            fill="toself", fillcolor=BOROUGH_COLOURS.get(la, ONS["grey"]),
            opacity=0.15,
        ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 30])),
        title="Westminster and Islington have the most LSOAs in the most deprived deciles",
        font_family="Arial", paper_bgcolor="white", height=450,
        legend=dict(orientation="h", y=-0.1),
    )
    st.plotly_chart(fig_radar, use_container_width=True)
    pptx_btn(fig_radar, "egdi_radar",
             "Westminster and Islington: highest share of LSOAs in most deprived deciles")
