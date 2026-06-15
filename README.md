# Westminster Children's Demographics Dashboard

> An interactive Streamlit dashboard exploring child poverty, ethnic demographics, Key Stage 4 attainment, and deprivation across Westminster LSOAs — benchmarked against CIPFA statistical neighbours.

Built for Westminster City Council analysts and policy teams. All charts are exportable to PowerPoint for stakeholder presentations.

---

## 🚀 Quick start

```bash
git clone https://github.com/<your-org>/westminster-children-dashboard.git
cd westminster-children-dashboard
pip install -r requirements.txt
streamlit run app.py
```

Then place the data files (see [Data sources](#-data-sources) below) in the `data/` folder and refresh.

---

## 📊 What this app shows

Westminster is home to significant child poverty alongside extreme wealth — **26.3% of children live in relative low income families (AHC, FYE 2025)**, with Church Street and Westbourne wards exceeding 44%. Benchmarked against Westminster's six CIPFA statistical neighbours (Kensington & Chelsea, Camden, Hammersmith & Fulham, Islington, and Wandsworth), the dashboard reveals where Westminster sits within its peer group and where within the borough inequality concentrates most sharply.

The app draws on five datasets spanning Census 2021 LSOA-level demographics, DWP child poverty statistics, DfE Key Stage 4 attainment by ethnic group, and the Ethnic Group Deprivation Index — giving a joined-up picture of children's outcomes across the borough.

**Audience:** Westminster City Council policy, children's services, public health, and scrutiny teams.

---

## 🗂️ Data sources

| Dataset | Source | Provider | Geography | Year |
|---------|--------|----------|-----------|------|
| [Children in Low Income Families (AHC Relative)](https://www.gov.uk/government/statistics/children-in-low-income-families-local-area-statistics-2022-to-2025) | DWP / HMRC | gov.uk | Local Authority & Ward | FYE 2024–2025 |
| [Key Stage 4 attainment by ethnicity](https://explore-education-statistics.service.gov.uk/data-tables/fast-track/1f770076-112b-45c2-5468-08de072d13df) | Department for Education | Explore Education Statistics | Local Authority (Inner London) | 2018/19–2024/25 |
| [Ethnic Group Deprivation Index (EGDI)](https://gedi.ac.uk/egdi/) | Lloyd et al. 2023 | GEDI / gedi.ac.uk | Local Authority | 2019 IMD-based |
| [RM006 — Age of youngest dependent child by household type](https://www.nomisweb.co.uk/) | Census 2021 | ONS Nomis | Westminster LSOAs | 2021 |
| [RM12 — Dependent children by ethnic group of HRP](https://www.nomisweb.co.uk/) | Census 2021 | ONS Nomis | Westminster LSOAs | 2021 |
| [RM033 — Ethnic group of dependent child by sex](https://www.nomisweb.co.uk/) | Census 2021 | ONS Nomis | Westminster LSOAs | 2021 |
| Westminster LSOA boundaries | ONS Open Geography Portal | ONS | Westminster LSOAs (TopoJSON) | 2011 |
| Local Authority boundaries | ONS Open Geography Portal | ONS | UK LAs (TopoJSON) | 2017 |

**CIPFA statistical neighbours used:** Westminster · Kensington & Chelsea · Camden · Hammersmith & Fulham · Islington · Wandsworth.
Source: [Trust for London CIPFA nearest neighbours](https://trustforlondon.org.uk/data/information-on-cipfa-nearest-statistical-neighbours/)

---

## 📁 Repository structure

```
westminster-children-dashboard/
├── app.py                          # Main Streamlit application
├── requirements.txt                # Pinned Python dependencies
├── README.md                       # This file
├── .gitignore                      # Excludes data files and secrets
├── .streamlit/
│   └── config.toml                 # ONS-branded theme + server config
└── data/                           # Place data files here (not committed to git)
    ├── children-in-low-income-families-local-area-statistics-2022-2025.ods
    ├── data-key-stage-4-performance.ods
    ├── data-key-stage-4-performance__1_.ods
    ├── EGDI-Local-Authority-profiles.xlsx
    ├── RM006_age_of_youngest_dependent_child_by_household_type.xlsx
    ├── RM12_dependent_children_by_ethnic_group_of_HRP.xlsx
    ├── RM033_ethic_group_dependent_child_by_sex.xlsx
    ├── LSOA_WCC__1_.json           # Westminster LSOA boundaries (TopoJSON)
    └── Local_Authority_UK.json     # UK LA boundaries (TopoJSON)
```

---

## 🖥️ App tabs

| Tab | What it shows |
|-----|---------------|
| **📍 Child Poverty — Borough Comparison** | Westminster vs CIPFA neighbours (LA level), ward-level breakdown within Westminster |
| **🗺️ Westminster LSOA Maps** | Choropleth maps of 128 Westminster LSOAs — switchable between 7 demographic metrics |
| **📚 KS4 Attainment** | Attainment 8 by ethnic group (2024/25) + trend 2018/19–2024/25 |
| **👥 Ethnicity & Demographics** | Ethnic group breakdown and age structure of children (Census 2021) |
| **⚖️ Ethnic Deprivation (EGDI)** | EGDI classification and deprivation decile distribution across neighbours |

---

## ⚙️ Configuration

### Data files

Place all data files in the `data/` folder. The app looks for files there first, then falls back to `/mnt/user-data/uploads/` (Claude.ai environment).

The `data/` folder is excluded from git (see `.gitignore`) — data files are too large and some may be restricted. Download them directly from the sources listed above.

### Streamlit theme

The app uses an ONS-branded theme defined in `.streamlit/config.toml`. No changes needed to run locally.

### Secrets

This app requires no API keys or secrets. No `.env` file is needed.

---

## 📦 Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `streamlit` | 1.45.1 | App framework |
| `plotly` | 6.1.2 | Interactive charts |
| `pandas` | 2.2.3 | Data wrangling |
| `python-pptx` | 1.0.2 | Per-chart PPTX slide export |
| `openpyxl` | 3.1.5 | Reading `.xlsx` files |
| `odfpy` | 1.4.1 | Reading `.ods` files |
| `topojson` | 1.9 | Converting TopoJSON boundaries to GeoJSON for maps |
| `kaleido` | 0.2.1 | Rendering Plotly charts to PNG for PPTX export |
| `numpy` | 2.2.6 | Numerical operations |

---

## 📤 Exporting charts to PowerPoint

Every chart in the app has a **⬇ Download slide (PPTX)** button directly beneath it. Clicking it generates a single-slide 16:9 PowerPoint with the chart rendered at high resolution (2400×1400px).

The sidebar also has an **📥 Export full deck** button that packages all currently visible charts into a single `.pptx` file — one slide per chart — ready to drop into a stakeholder presentation.

---

## 🗺️ Data notes

- **Child poverty figures** are after-housing-costs (AHC) relative poverty for children aged 0–15. AHC is generally considered the more meaningful measure as it accounts for housing cost variation across London.
- **KS4 data** covers state-funded schools only. Small pupil numbers at ethnic group × borough level lead to suppressed cells (shown as "no data"). 2020/21 and 2021/22 figures should be interpreted with caution due to COVID-19 assessment disruption.
- **Census 2021 LSOA data** uses 2011 LSOA boundaries. Westminster has 128 LSOAs across 18 electoral wards.
- **EGDI** is based on IMD 2019 income deprivation domain. "More ethnic inequality" means the distribution of deprivation across ethnic groups within the borough is significantly unequal — not that the borough is more deprived overall.
- **CIPFA neighbours** are Westminster's six nearest statistical neighbours as defined by CIPFA's 2022 methodology, sourced from Trust for London.

---

## 📜 Licence

Code: [MIT Licence](LICENSE)

Data licences vary by source — see individual dataset pages. ONS data is available under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/). DWP and DfE data are similarly OGL-licensed. EGDI is available for research and policy use — see [gedi.ac.uk](https://gedi.ac.uk/egdi/) for terms.
