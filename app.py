from io import BytesIO
import html
import re

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Strategic District Field Guide",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# STYLING
# ============================================================

CUSTOM_CSS = """
<style>
:root {
    --navy:#102a43;
    --blue:#1d4ed8;
    --blue-soft:#eff6ff;
    --blue-border:#bfdbfe;

    --teal:#0f766e;
    --teal-soft:#ecfdf5;
    --teal-border:#99f6e4;

    --purple:#6d28d9;
    --purple-soft:#f5f3ff;
    --purple-border:#ddd6fe;

    --amber:#92400e;
    --amber-soft:#fffbeb;
    --amber-border:#fde68a;

    --green:#166534;
    --green-soft:#f0fdf4;
    --green-border:#bbf7d0;

    --red:#991b1b;
    --red-soft:#fef2f2;
    --red-border:#fecaca;

    --slate:#334155;
    --slate-soft:#f8fafc;
    --slate-border:#e2e8f0;

    --page:#f3f6fa;
    --surface:#ffffff;
    --surface-2:#f8fafc;
    --text:#0f172a;
    --muted:#475569;
    --border:#cbd5e1;
}

/* Force readable light-mode app even if browser/Streamlit is in dark mode */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stHeader"] {
    background: var(--page) !important;
    color: var(--text) !important;
}

html,
body,
p,
li,
span,
div,
label,
section,
article,
.stMarkdown,
.stMarkdown p,
.stMarkdown li,
.stMarkdown span,
.stMarkdown div {
    color: var(--text) !important;
}

/* Main content width */
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2.5rem;
    max-width: 1180px;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid var(--border) !important;
}

[data-testid="stSidebar"] * {
    color: var(--text) !important;
}

[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] small {
    color: var(--muted) !important;
}

/* Headings */
h1, h2, h3, h4, h5, h6 {
    color: var(--navy) !important;
}

.main-title {
    font-size: 2rem;
    font-weight: 850;
    color: var(--navy) !important;
    margin-bottom: .25rem;
    letter-spacing:-.02em;
}

.subtitle {
    color: var(--muted) !important;
    font-size: .98rem;
    margin-bottom: 1rem;
}

.section-title {
    color: var(--navy) !important;
    font-size: 1.25rem;
    font-weight: 800;
    margin: .6rem 0 .3rem;
}

.helper-text {
    color: var(--muted) !important;
    font-size: .9rem;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 14px !important;
    padding: .25rem !important;
}

.stTabs [data-baseweb="tab"] {
    color: var(--slate) !important;
    background: transparent !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
}

.stTabs [aria-selected="true"] {
    background: var(--blue-soft) !important;
    color: var(--blue) !important;
}

/* Metrics */
[data-testid="stMetric"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 16px !important;
    padding: .85rem !important;
    box-shadow: 0 4px 14px rgba(15, 23, 42, .05) !important;
}

[data-testid="stMetric"] * {
    color: var(--text) !important;
}

[data-testid="stMetricLabel"] {
    color: var(--muted) !important;
}

[data-testid="stMetricValue"] {
    color: var(--navy) !important;
}

/* Inputs */
input,
textarea,
select,
[data-baseweb="input"],
[data-baseweb="select"],
[data-baseweb="textarea"] {
    background: #ffffff !important;
    color: var(--text) !important;
    border-color: var(--border) !important;
}

input::placeholder,
textarea::placeholder {
    color: #64748b !important;
}

/* Buttons */
.stButton button,
.stDownloadButton button {
    background: var(--blue) !important;
    color: #ffffff !important;
    border: 1px solid var(--blue) !important;
    border-radius: 10px !important;
    font-weight: 750 !important;
}

.stButton button:hover,
.stDownloadButton button:hover {
    background: #1e40af !important;
    border-color: #1e40af !important;
    color: #ffffff !important;
}

/* Expanders */
.streamlit-expanderHeader,
[data-testid="stExpander"] summary {
    color: var(--navy) !important;
    font-weight: 750 !important;
}

[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 14px !important;
}

[data-testid="stExpander"] * {
    color: var(--text) !important;
}

/* Dataframes */
[data-testid="stDataFrame"] {
    background: #ffffff !important;
    color: var(--text) !important;
}

/* Cards */
.card {
    border:1px solid var(--border);
    border-radius:20px;
    padding:1.05rem;
    margin-bottom:1rem;
    background:var(--surface);
    color: var(--text) !important;
    box-shadow:0 8px 24px rgba(15,23,42,.08);
}

.card,
.card * {
    color: var(--text) !important;
}

.card h3 {
    color:var(--navy) !important;
    margin-top:0;
    margin-bottom:.2rem;
    font-size:1.35rem;
    letter-spacing:-.01em;
}

.meta {
    color:var(--muted) !important;
    font-size:.88rem;
    margin-bottom:.75rem;
    line-height:1.45;
}

.lead {
    background:var(--blue-soft);
    border-left:4px solid var(--blue);
    border-radius:14px;
    padding:.75rem;
    margin:.6rem 0;
    line-height:1.45;
    color: var(--text) !important;
}

.lead * {
    color: var(--text) !important;
}

/* Quick prep */
.quickprep {
    background: #ffffff;
    border:1px solid var(--blue-border);
    border-left:5px solid var(--blue);
    border-radius:16px;
    padding:.9rem;
    margin:.7rem 0 .8rem;
    color: var(--text) !important;
    box-shadow: 0 4px 16px rgba(15, 23, 42, .05);
}

.quickprep,
.quickprep * {
    color: var(--text) !important;
}

.quickprep-title {
    color:var(--navy) !important;
    font-weight:850;
    font-size:.95rem;
    text-transform:uppercase;
    letter-spacing:.04em;
    margin-bottom:.25rem;
}

/* Badges and chips */
.badge,
.chip {
    display:inline-block;
    border-radius:999px;
    padding:.25rem .58rem;
    margin:.14rem .16rem .14rem 0;
    font-size:.76rem;
    font-weight:750;
    line-height:1.2;
    border:1px solid transparent;
}

.chip {
    background:var(--slate-soft);
    color:var(--slate) !important;
    border-color:var(--slate-border);
}

.tag-math {
    background:#dbeafe;
    color:#1e3a8a !important;
    border-color:#bfdbfe;
}

.tag-mtss {
    background:#ccfbf1;
    color:#134e4a !important;
    border-color:#99f6e4;
}

.tag-spedell {
    background:#ede9fe;
    color:#4c1d95 !important;
    border-color:#ddd6fe;
}

.tag-ccmr {
    background:#fef3c7;
    color:#78350f !important;
    border-color:#fde68a;
}

.tag-teacher {
    background:#f1f5f9;
    color:#1e293b !important;
    border-color:#e2e8f0;
}

.tag-hqim {
    background:#e0e7ff;
    color:#312e81 !important;
    border-color:#c7d2fe;
}

.tag-funding {
    background:#fef9c3;
    color:#713f12 !important;
    border-color:#fde68a;
}

.tag-relationship {
    background:#dcfce7;
    color:#14532d !important;
    border-color:#bbf7d0;
}

.tag-default {
    background:#eef2ff;
    color:#312e81 !important;
    border-color:#e0e7ff;
}

/* Priority pills */
.priority {
    display:inline-block;
    border-radius:999px;
    padding:.28rem .62rem;
    font-size:.75rem;
    font-weight:850;
    vertical-align:middle;
}

.priority-very-high {
    background:var(--red-soft);
    color:var(--red) !important;
    border: 1px solid var(--red-border);
}

.priority-high {
    background:var(--green-soft);
    color:var(--green) !important;
    border: 1px solid var(--green-border);
}

.priority-medium-high {
    background:var(--amber-soft);
    color:var(--amber) !important;
    border: 1px solid var(--amber-border);
}

.priority-medium {
    background:#e0f2fe;
    color:#075985 !important;
    border: 1px solid #bae6fd;
}

.metric-note {
    color:var(--muted) !important;
    font-size:.84rem;
    margin-top:-.25rem;
}

.howto-box {
    background:#ffffff;
    color: var(--text) !important;
    border:1px solid var(--border);
    border-radius:18px;
    padding:1rem 1.1rem;
    margin:.75rem 0;
    box-shadow:0 4px 16px rgba(15,23,42,.05);
}

.howto-box,
.howto-box * {
    color: var(--text) !important;
}

hr {
    border:0;
    border-top:1px solid var(--border);
    margin:1.2rem 0;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ============================================================
# EXPECTED LARGER WORKBOOK SHEETS
# ============================================================

SHEET_STRATEGIC = "Strategic Indicators"
SHEET_SCORECARD = "Master Scorecard"
SHEET_BASIC = "Basic District Info"
SHEET_CONTACTS = "District Contacts"
SHEET_LEADERSHIP = "Leadership and Governance"
SHEET_CSI = "CSI"
SHEET_TSI = "TSI"
DISTRICT_COL = "District Name"


# ============================================================
# BASIC HELPERS
# ============================================================

def safe_html(value):
    return html.escape(str(value))


def normalize_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_key(value):
    text = normalize_text(value).lower()
    text = text.replace("\n", " ").replace("\r", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def district_key(value):
    return normalize_text(value).upper().strip()


def make_unique_columns(columns):
    seen = {}
    output = []
    for col in columns:
        base = normalize_text(col) or "Unnamed"
        if base not in seen:
            seen[base] = 0
            output.append(base)
        else:
            seen[base] += 1
            output.append(f"{base}.{seen[base]}")
    return output


def clean_columns(df):
    if df.empty:
        return df
    df = df.copy()
    df.columns = make_unique_columns([str(c).strip().replace("\n", " ").replace("\r", " ") for c in df.columns])
    drop_cols = [c for c in df.columns if normalize_key(c).startswith("methods:") or normalize_key(c).startswith("unnamed")]
    if drop_cols:
        df = df.drop(columns=drop_cols)
    return df


def find_sheet_name(xls, desired_name):
    lookup = {normalize_key(sheet): sheet for sheet in xls.sheet_names}
    return lookup.get(normalize_key(desired_name))


def read_table_sheet(xls, desired_sheet_name):
    actual_sheet = find_sheet_name(xls, desired_sheet_name)
    if not actual_sheet:
        return pd.DataFrame()

    raw = pd.read_excel(xls, sheet_name=actual_sheet, header=None, engine="openpyxl")
    if raw.empty:
        return pd.DataFrame()

    header_row = None
    for idx, row in raw.iterrows():
        values = [normalize_text(v) for v in row.tolist()]
        if DISTRICT_COL in values:
            header_row = idx
            break

    if header_row is None:
        return pd.DataFrame()

    headers = [normalize_text(v) for v in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    data.columns = headers
    data = clean_columns(data)
    data = data.dropna(how="all")

    if DISTRICT_COL in data.columns:
        data = data[data[DISTRICT_COL].notna()]
        data = data[data[DISTRICT_COL].astype(str).str.strip() != ""]

    return data.reset_index(drop=True)


def find_col(df, exact_names=None, contains_all=None):
    if df.empty:
        return None
    exact_names = exact_names or []
    contains_all = contains_all or []
    normalized_map = {normalize_key(col): col for col in df.columns}

    for name in exact_names:
        key = normalize_key(name)
        if key in normalized_map:
            return normalized_map[key]

    if contains_all:
        for col in df.columns:
            col_key = normalize_key(col)
            if all(term.lower() in col_key for term in contains_all):
                return col

    return None


def get_value(row_or_dict, col_name, default=""):
    if row_or_dict is None:
        return default

    if isinstance(row_or_dict, dict):
        for key, value in row_or_dict.items():
            if normalize_key(key) == normalize_key(col_name):
                return value
        return default

    try:
        for key in row_or_dict.index:
            if normalize_key(key) == normalize_key(col_name):
                return row_or_dict.get(key, default)
    except Exception:
        return default

    return default


def format_number(value, decimals=2):
    if normalize_text(value) == "":
        return ""
    try:
        return f"{float(value):.{decimals}f}"
    except Exception:
        return normalize_text(value)


def format_enrollment(value):
    if normalize_text(value) == "":
        return ""
    try:
        return f"{int(float(value)):,}"
    except Exception:
        return normalize_text(value)


def format_percent(value, decimals=0):
    if normalize_text(value) == "":
        return ""
    try:
        number = float(value)
        pct = number * 100 if abs(number) <= 1 else number
        if decimals == 0:
            return f"{pct:.0f}%"
        return f"{pct:.{decimals}f}%"
    except Exception:
        return normalize_text(value)


def tag_class(tag):
    mapping = {
        "Math": "tag-math",
        "MTSS": "tag-mtss",
        "SPED/ELL": "tag-spedell",
        "CCMR": "tag-ccmr",
        "Teacher Capacity": "tag-teacher",
        "Curriculum / HQIM": "tag-hqim",
        "Funding / Grants": "tag-funding",
        "Existing Relationship": "tag-relationship",
    }
    return mapping.get(tag, "tag-default")


def badge_html(tag):
    return f'<span class="badge {tag_class(tag)}">{safe_html(tag)}</span>'


def chip_html(item):
    return f'<span class="chip">{safe_html(item)}</span>'


# ============================================================
# COLUMN DETECTION
# ============================================================

def detect_scorecard_columns(scorecard_df):
    return {
        "district": find_col(scorecard_df, exact_names=["District Name"]),
        "enrollment": find_col(scorecard_df, exact_names=["Enrollment"]),
        "strategic_score": find_col(scorecard_df, exact_names=["Strategic Weighted Score"], contains_all=["strategic", "weighted", "score"]),
        "overall_score": find_col(scorecard_df, exact_names=[
            "Overall Weighted Score (Strategic+Contract+Relationship)",
            "Overall Weighted Score (Strategic+Relationship Weighted)",
            "Overall Weighted Score",
        ], contains_all=["overall", "weighted", "score"]),
        "tier": find_col(scorecard_df, exact_names=["Tier"]),
        "existing_contracts": find_col(scorecard_df, exact_names=["Existing Contracts in Region (Yes/No)"]),
        "existing_relationships": find_col(scorecard_df, exact_names=["Existing Relationships"]),
    }


# ============================================================
# LOOKUPS AND ELIGIBILITY
# ============================================================

def is_substantive_strategic_row(row):
    fields_to_check = [
        "Strategic Plan Themes", "Math Improvement Mentioned", "Math Priority Strength",
        "Intervention Focus", "Intervention Focus Details", "Teacher Capacity/PD Focus",
        "Teacher Capacity Details", "Career Readiness Mentioned", "Career Readiness Details",
        "SPED/ELL Improvement Mentioned", "SPED/ELL Details", "MTSS/Tiered Support Mentioned",
        "MTSS Details", "Curriculum Review/Adoption Activity", "Curriculum Details",
        "Active Grants (Yes/No)", "Grants Details", "Sources", "Notes",
    ]
    return sum(1 for field in fields_to_check if normalize_text(get_value(row, field))) >= 2


def build_score_lookup(scorecard_df, score_cols):
    lookup = {}
    district_col = score_cols.get("district")
    if scorecard_df.empty or not district_col:
        return lookup
    for _, row in scorecard_df.iterrows():
        district = normalize_text(row.get(district_col, ""))
        if district:
            lookup[district_key(district)] = row.to_dict()
    return lookup


def build_basic_lookup(basic_df):
    lookup = {}
    if basic_df.empty or DISTRICT_COL not in basic_df.columns:
        return lookup
    for _, row in basic_df.iterrows():
        district = normalize_text(row.get(DISTRICT_COL, ""))
        if district:
            lookup[district_key(district)] = row.to_dict()
    return lookup


def build_csi_tsi_counts(csi_df, tsi_df):
    csi_counts = {}
    tsi_counts = {}
    if not csi_df.empty and DISTRICT_COL in csi_df.columns:
        csi_counts = csi_df.groupby(csi_df[DISTRICT_COL].astype(str).str.upper().str.strip()).size().to_dict()
    if not tsi_df.empty and DISTRICT_COL in tsi_df.columns:
        tsi_counts = tsi_df.groupby(tsi_df[DISTRICT_COL].astype(str).str.upper().str.strip()).size().to_dict()
    return csi_counts, tsi_counts


def determine_priority(tier, score):
    if tier == "Tier 1":
        return "Very High"
    if tier == "Tier 2":
        return "High"
    try:
        if float(score) >= 3.5:
            return "Medium-High"
    except Exception:
        pass
    return "Medium"


# ============================================================
# CONTACTS
# ============================================================

def build_contacts_lookup_from_contacts(contacts_df):
    lookup = {}
    if contacts_df.empty or DISTRICT_COL not in contacts_df.columns:
        return lookup

    role_priority = [
        "SUPERINTENDENT", "CHIEF ACADEMIC OFFICER", "ASSISTANT SUPERINTENDENT",
        "CURRICULUM DIRECTOR", "MATH COORDINATOR", "CTE COLLEGE CAREER READINESS DIRECTOR",
        "SPECIAL EDUCATION DIRECTOR", "ESL ELL COORDINATOR", "DIRECTOR ASSESSMENT DATA",
        "PROFESSIONAL LEARNING DIRECTOR", "ELA LITERACY COORDINATOR",
    ]

    df = contacts_df.copy()
    if "Position" in df.columns:
        df["_priority"] = df["Position"].astype(str).str.upper().apply(lambda x: role_priority.index(x) if x in role_priority else 99)
        df = df.sort_values([DISTRICT_COL, "_priority"])

    for _, row in df.iterrows():
        district = normalize_text(row.get(DISTRICT_COL, ""))
        if not district:
            continue
        first = normalize_text(row.get("First Name", ""))
        last = normalize_text(row.get("Last Name", ""))
        title = normalize_text(row.get("Title", ""))
        position = normalize_text(row.get("Position", ""))
        name = f"{first} {last}".strip()

        if name and title:
            contact = f"{name} — {title}"
        elif name and position:
            contact = f"{name} — {position}"
        elif title:
            contact = title
        else:
            contact = ""

        if contact:
            lookup.setdefault(district_key(district), [])
            if contact not in lookup[district_key(district)]:
                lookup[district_key(district)].append(contact)

    return lookup


def build_contacts_lookup_from_leadership(leadership_df):
    lookup = {}
    if leadership_df.empty or DISTRICT_COL not in leadership_df.columns:
        return lookup

    for _, row in leadership_df.iterrows():
        district = normalize_text(row.get(DISTRICT_COL, ""))
        if not district:
            continue
        contacts = []
        superintendent = normalize_text(get_value(row, "Superintendent"))
        curriculum_lead = normalize_text(get_value(row, "Curriculum Lead"))
        cte_lead = normalize_text(get_value(row, "CTE Lead"))
        math_lead = normalize_text(get_value(row, "Math Lead"))

        if superintendent:
            contacts.append(f"{superintendent} — Superintendent")
        if curriculum_lead:
            contacts.append(f"{curriculum_lead} — Curriculum / Academic Lead")
        if cte_lead:
            contacts.append(f"{cte_lead} — CTE / Career Readiness Lead")
        if math_lead:
            contacts.append(f"{math_lead} — Math Lead")

        cleaned = []
        for contact in contacts:
            contact = contact.replace("[", "").replace("]", "")
            contact = contact.replace("(mailto:", " | ").replace(")", "")
            cleaned.append(contact)
        lookup[district_key(district)] = cleaned[:6]
    return lookup


def get_contacts(district_name, contacts_lookup, leadership_lookup):
    key = district_key(district_name)

    if key in contacts_lookup and contacts_lookup[key]:
        return contacts_lookup[key][:6]

    return leadership_lookup.get(key, [])[:6]


# ============================================================
# STRATEGIC CARD LANGUAGE
# ============================================================

def infer_tags(strategic_row, score_row):
    tags = []
    combined_text = " ".join(normalize_text(v) for v in strategic_row.to_dict().values()).lower()

    if normalize_text(get_value(strategic_row, "Math Improvement Mentioned")) or normalize_text(get_value(strategic_row, "Math Priority Strength")):
        tags.append("Math")
    if normalize_text(get_value(strategic_row, "Intervention Focus")) or normalize_text(get_value(strategic_row, "MTSS/Tiered Support Mentioned")):
        tags.append("MTSS")
    if ("sped" in combined_text or "special education" in combined_text or "english learner" in combined_text or "emergent bilingual" in combined_text or "ell" in combined_text or "multilingual" in combined_text):
        tags.append("SPED/ELL")
    if normalize_text(get_value(strategic_row, "Career Readiness Mentioned")) or normalize_text(get_value(strategic_row, "Career Readiness Details")):
        tags.append("CCMR")
    if normalize_text(get_value(strategic_row, "Teacher Capacity/PD Focus")) or normalize_text(get_value(strategic_row, "Teacher Capacity Details")):
        tags.append("Teacher Capacity")
    if normalize_text(get_value(strategic_row, "Curriculum Review/Adoption Activity")) or normalize_text(get_value(strategic_row, "Curriculum Details")):
        tags.append("Curriculum / HQIM")
    if normalize_text(get_value(strategic_row, "Active Grants (Yes/No)")) or normalize_text(get_value(strategic_row, "Grants Details")):
        tags.append("Funding / Grants")
    if normalize_text(get_value(score_row, "Existing Relationships")).lower() == "yes":
        tags.append("Existing Relationship")

    if not tags:
        tags.append("Strategic Review")

    unique = []
    for tag in tags:
        if tag not in unique:
            unique.append(tag)
    return unique


def build_alignment(tags):
    alignment = []
    if "Math" in tags:
        alignment.append("Elevation Station Math Games — K–8 math practice, fluency, reinforcement, and engagement.")
        alignment.append("Elevation intervention curriculum — targeted K–5 support where Tier 2/Tier 3 math or foundational gaps are present.")
    if "MTSS" in tags:
        alignment.append("PCG MTSS consulting — system design, campus implementation, intervention fidelity, and data-to-action routines.")
    if "SPED/ELL" in tags:
        alignment.append("PCG SPED / multilingual learner support — inclusive practice, service delivery, access to grade-level instruction, and subgroup progress monitoring.")
    if "CCMR" in tags:
        alignment.append("RISE Career & Math Mini Lessons — grades 6–9 career-connected math, pathway awareness, and applied readiness.")
    if "Curriculum / HQIM" in tags:
        alignment.append("PCG curriculum/HQIM implementation support — adoption fidelity, coaching, PLC routines, and change management.")
    if "Teacher Capacity" in tags:
        alignment.append("PCG professional learning and instructional implementation support — teacher capacity, coaching, data use, and scalable instructional routines.")
    if "Funding / Grants" in tags:
        alignment.append("Funding alignment support — connect implementation supports to existing grant, Title, or strategic funding streams where appropriate.")
    if not alignment:
        alignment.append("Discovery needed — validate strategic needs, stakeholder priorities, and fit before positioning specific resources.")
    return alignment


def build_lead_with(card):
    tags = card.get("tags", [])
    district_name = card.get("name", "the district")
    if "Existing Relationship" in tags and card.get("tier") == "Tier 1":
        relationship_phrase = "given the existing relationship and strong strategic fit"
    elif "Existing Relationship" in tags:
        relationship_phrase = "building from the existing relationship"
    else:
        relationship_phrase = "as an initial consultative entry point"
    top_tags = [tag for tag in tags if tag not in ["Existing Relationship", "Funding / Grants"]]
    top_tags_text = ", ".join(top_tags[:4]).lower() or "strategic implementation support"
    return f"Lead with {top_tags_text} {relationship_phrase}. Frame the conversation around helping {district_name} move from stated priorities to consistent campus-level execution without adding unnecessary burden."


def build_questions(card):
    tags = card.get("tags", [])
    questions = []
    if "Math" in tags:
        questions.append("How are math goals being translated into weekly instructional decisions, student practice routines, and progress monitoring?")
    else:
        questions.append("How are the district’s strategic priorities being translated into consistent campus-level routines?")
    if "MTSS" in tags:
        questions.append("When students are identified for additional support, where does the process tend to slow down — grouping, scheduling, intervention materials, progress monitoring, or teacher capacity?")
    else:
        questions.append("Where do you see the biggest gap between the strategic plan and day-to-day classroom execution?")
    if "SPED/ELL" in tags:
        questions.append("Where do students with disabilities or emergent bilingual students most often lose access to grade-level expectations?")
    if "Teacher Capacity" in tags or "Curriculum / HQIM" in tags:
        questions.append("After initial training or rollout, where do teachers tend to need the most help — planning, pacing, differentiation, student practice, or responding to data?")
    if "CCMR" in tags:
        questions.append("How early are students connecting academic skills, especially math, to future pathways and readiness expectations?")
    questions.append("If current implementation barriers remain, what are the implications for students, staff, accountability outcomes, and community confidence?")
    questions.append("What would make an external partner feel like implementation support rather than another initiative?")
    questions.append("If you were to pilot targeted support, which campuses, grade bands, or student groups would create the clearest proof point?")
    return questions


def build_listen_for(tags):
    tag_map = {
        "Math": ["math growth", "early numeracy", "Algebra readiness", "STAAR math", "student practice"],
        "MTSS": ["intervention fidelity", "Tier 2", "Tier 3", "progress monitoring", "campus variation"],
        "SPED/ELL": ["access to grade-level instruction", "service delivery", "emergent bilingual students", "students with disabilities", "subgroup gaps"],
        "CCMR": ["pathways", "industry-based certifications", "TSIA2", "dual credit", "career awareness"],
        "Teacher Capacity": ["teacher burden", "coaching", "professional learning", "PLC routines", "instructional consistency"],
        "Curriculum / HQIM": ["adoption fidelity", "HQIM", "curriculum implementation", "Eureka", "Bluebonnet", "instructional materials"],
        "Funding / Grants": ["Title funding", "grant alignment", "LASSO", "federal funds", "implementation funding"],
        "Existing Relationship": ["existing relationship", "current contract", "expansion", "trusted partner"],
    }
    listen_for = []
    for tag in tags:
        listen_for.extend(tag_map.get(tag, []))
    listen_for.extend(["data cycles", "implementation barriers", "capacity constraints"])
    unique = []
    for item in listen_for:
        if item not in unique:
            unique.append(item)
    return unique


def build_avoid(tags):
    avoid = ["Do not lead with a product pitch.", "Avoid positioning support as a replacement for the district’s current strategy or adopted curriculum."]
    if "Curriculum / HQIM" in tags:
        avoid.append("Avoid implying the district needs another curriculum; frame support around implementation, practice, and adoption fidelity.")
    if "SPED/ELL" in tags:
        avoid.append("Avoid treating subgroup performance as a side issue; connect it to access, instruction, service delivery, and progress monitoring.")
    if "Existing Relationship" in tags:
        avoid.append("Avoid ignoring the existing relationship context; build from credibility and continuity.")
    return avoid


# ============================================================
# CARD BUILDING
# ============================================================

def build_card(strategic_row, score_row, basic_lookup, contacts_lookup, leadership_lookup, csi_counts, tsi_counts, score_cols):
    district_name = normalize_text(get_value(strategic_row, DISTRICT_COL))
    key = district_key(district_name)

    raw_score = score_row.get(score_cols["overall_score"], "")
    score = format_number(raw_score, 2)
    strategic_score = format_number(score_row.get(score_cols["strategic_score"], ""), 2) if score_cols.get("strategic_score") else ""
    tier = normalize_text(score_row.get(score_cols["tier"], ""))
    enrollment = format_enrollment(score_row.get(score_cols["enrollment"], "")) if score_cols.get("enrollment") else ""

    tags = infer_tags(strategic_row, score_row)
    priority = determine_priority(tier, raw_score)
    basic = basic_lookup.get(key, {})

    signals = []
    themes = normalize_text(get_value(strategic_row, "Strategic Plan Themes"))
    math_strength = normalize_text(get_value(strategic_row, "Math Priority Strength"))
    intervention_details = normalize_text(get_value(strategic_row, "Intervention Focus Details"))
    sped_ell_details = normalize_text(get_value(strategic_row, "SPED/ELL Details"))
    teacher_details = normalize_text(get_value(strategic_row, "Teacher Capacity Details"))
    career_details = normalize_text(get_value(strategic_row, "Career Readiness Details"))
    mtss_details = normalize_text(get_value(strategic_row, "MTSS Details"))
    curriculum_details = normalize_text(get_value(strategic_row, "Curriculum Details"))
    grants_details = normalize_text(get_value(strategic_row, "Grants Details"))

    if themes:
        signals.append(f"Strategic themes: {themes}")
    if math_strength:
        signals.append(f"Math signal: {math_strength}")

    csi = normalize_text(get_value(basic, "CSI Schools")) or str(csi_counts.get(key, ""))
    tsi = normalize_text(get_value(basic, "TSI Schools")) or str(tsi_counts.get(key, ""))
    if csi or tsi:
        signals.append(f"Accountability pressure: {csi or '0'} CSI schools and {tsi or '0'} TSI schools.")

    context_parts = []
    context_fields = [
        ("schools", "Number of Schools", "text"),
        ("grade span", "Grade Span Served", "text"),
        ("setting", "Urban/Suburban/Rural", "text"),
        ("economically disadvantaged", "% Economically Disadvantaged", "percent"),
        ("English learner", "% English Learner", "percent"),
        ("special education", "% Special Education", "percent"),
        ("major student group", "Major Student Groups", "text"),
        ("student growth trend", "Student Growth Trend", "percent1"),
    ]
    for label, col, kind in context_fields:
        raw_value = get_value(basic, col)
        value = normalize_text(raw_value)
        if value:
            if kind == "percent":
                value = format_percent(raw_value, 0)
            elif kind == "percent1":
                value = format_percent(raw_value, 1)
            context_parts.append(f"{label}: {value}")
    if context_parts:
        signals.append("District context: " + "; ".join(context_parts) + ".")

    if intervention_details:
        signals.append(f"Intervention signal: {intervention_details}")
    if mtss_details:
        signals.append(f"MTSS signal: {mtss_details}")
    if sped_ell_details:
        signals.append(f"SPED/ELL signal: {sped_ell_details}")
    if teacher_details:
        signals.append(f"Teacher capacity signal: {teacher_details}")
    if career_details:
        signals.append(f"CCMR / career readiness signal: {career_details}")
    if curriculum_details:
        signals.append(f"Curriculum / implementation signal: {curriculum_details}")
    if grants_details:
        signals.append(f"Funding / grants signal: {grants_details}")

    existing_relationships = normalize_text(get_value(score_row, "Existing Relationships"))
    existing_contracts = normalize_text(get_value(score_row, "Existing Contracts in Region (Yes/No)"))
    if existing_contracts and existing_relationships:
        signals.append(f"Relationship context: existing contracts in region = {existing_contracts}; existing relationships = {existing_relationships}.")
    elif existing_relationships:
        signals.append(f"Relationship context: existing relationships = {existing_relationships}.")
    elif existing_contracts:
        signals.append(f"Relationship context: existing contracts in region = {existing_contracts}.")

    card = {
        "name": district_name,
        "tier": tier,
        "score": score,
        "strategic_score": strategic_score,
        "enrollment": enrollment,
        "priority": priority,
        "tags": tags,
        "contacts": get_contacts(district_name, contacts_lookup, leadership_lookup),
        "signals": signals,
        "alignment": build_alignment(tags),
        "questions": [],
        "listen": [],
        "avoid": [],
    }
    card["lead"] = build_lead_with(card)
    card["questions"] = build_questions(card)
    card["listen"] = build_listen_for(tags)
    card["avoid"] = build_avoid(tags)
    return card


# ============================================================
# LOAD WORKBOOK
# ============================================================

def load_cards_from_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

    strategic_df = read_table_sheet(xls, SHEET_STRATEGIC)
    scorecard_df = read_table_sheet(xls, SHEET_SCORECARD)
    basic_df = read_table_sheet(xls, SHEET_BASIC)
    contacts_df = read_table_sheet(xls, SHEET_CONTACTS)
    leadership_df = read_table_sheet(xls, SHEET_LEADERSHIP)
    csi_df = read_table_sheet(xls, SHEET_CSI)
    tsi_df = read_table_sheet(xls, SHEET_TSI)

    if strategic_df.empty or scorecard_df.empty:
        return [], xls, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, pd.DataFrame(), {}

    score_cols = detect_scorecard_columns(scorecard_df)
    score_lookup = build_score_lookup(scorecard_df, score_cols)
    basic_lookup = build_basic_lookup(basic_df)
    contacts_lookup = build_contacts_lookup_from_contacts(contacts_df)
    leadership_lookup = build_contacts_lookup_from_leadership(leadership_df)
    csi_counts, tsi_counts = build_csi_tsi_counts(csi_df, tsi_df)

    cards = []
    audit_rows = []

    for _, strategic_row in strategic_df.iterrows():
        district_name = normalize_text(get_value(strategic_row, DISTRICT_COL))
        if not district_name:
            continue

        key = district_key(district_name)
        substantive = is_substantive_strategic_row(strategic_row)
        score_row = score_lookup.get(key, {})
        matched = bool(score_row)
        tier = normalize_text(score_row.get(score_cols.get("tier"), "")) if matched and score_cols.get("tier") else ""
        overall = score_row.get(score_cols.get("overall_score"), "") if matched and score_cols.get("overall_score") else ""
        eligible = substantive and matched and bool(tier) and bool(normalize_text(overall))

        audit_rows.append({
            "District": district_name,
            "Substantive Strategic Indicators": substantive,
            "Matched Master Scorecard": matched,
            "Detected Tier Column": score_cols.get("tier"),
            "Tier": tier,
            "Detected Overall Score Column": score_cols.get("overall_score"),
            "Overall Score": normalize_text(overall),
            "Eligible": eligible,
        })

        if eligible:
            cards.append(build_card(strategic_row, score_row, basic_lookup, contacts_lookup, leadership_lookup, csi_counts, tsi_counts, score_cols))

    audit_df = pd.DataFrame(audit_rows)
    return cards, xls, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, audit_df, score_cols


# ============================================================
# DEBUG
# ============================================================

def show_workbook_debug(xls, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, audit_df, score_cols):
    st.markdown('<div class="section-title">Workbook Diagnostics</div>', unsafe_allow_html=True)
    st.markdown('<div class="helper-text">Use this tab to verify that the uploaded workbook was read correctly.</div>', unsafe_allow_html=True)
    st.markdown("### Sheets found")
    st.write(xls.sheet_names)
    st.markdown("### Detected scorecard columns")
    st.write(score_cols)
    for label, df in [
        ("Strategic Indicators", strategic_df),
        ("Master Scorecard", scorecard_df),
        ("Basic District Info", basic_df),
        ("District Contacts", contacts_df),
        ("Leadership and Governance", leadership_df),
        ("CSI", csi_df),
        ("TSI", tsi_df),
    ]:
        with st.expander(label, expanded=False):
            st.write(f"Rows: {len(df)}")
            st.write("Columns:")
            st.write(list(df.columns))
            if DISTRICT_COL in df.columns:
                st.write("First district names found:")
                st.write(df[DISTRICT_COL].dropna().astype(str).head(15).tolist())
    st.markdown("### Eligibility audit")
    st.dataframe(audit_df, use_container_width=True)


# ============================================================
# SEARCH / FILTER
# ============================================================

def search_blob(card):
    values = [card.get("name", ""), card.get("tier", ""), str(card.get("score", "")), str(card.get("strategic_score", "")), card.get("enrollment", ""), card.get("priority", ""), card.get("lead", "")]
    for key in ["tags", "signals", "alignment", "contacts", "questions", "listen", "avoid"]:
        values.extend(card.get(key, []))
    return " ".join(map(str, values)).lower()


def filter_cards(cards, query, selected_tiers, selected_tags, shortlist, show_shortlist_only):
    query = query.strip().lower()
    filtered = []
    for card in cards:
        if selected_tiers and card.get("tier") not in selected_tiers:
            continue
        if selected_tags and not any(tag in card.get("tags", []) for tag in selected_tags):
            continue
        if show_shortlist_only and card.get("name") not in shortlist:
            continue
        if query and query not in search_blob(card):
            continue
        filtered.append(card)
    return filtered


# ============================================================
# RENDERING
# ============================================================

def render_quick_prep(card):
    top_listen = card.get("listen", [])[:10]
    best_question = card.get("questions", [""])[0] if card.get("questions") else ""
    avoid = card.get("avoid", [""])[0] if card.get("avoid") else ""
    chips = "".join(chip_html(item) for item in top_listen)
    st.markdown(f"""
    <div class="quickprep">
        <div class="quickprep-title">Quick Prep</div>
        <strong>Best opening question:</strong> {safe_html(best_question)}<br>
        <strong>Listen for:</strong><br>{chips}<br>
        <strong>Avoid:</strong> {safe_html(avoid)}
    </div>
    """, unsafe_allow_html=True)


def render_card(card, conference_mode=False):
    priority_class = card.get("priority", "Medium").lower().replace(" ", "-")
    badges = "".join(badge_html(tag) for tag in card.get("tags", []))

    st.markdown(f"""
        <div class="card">
            <h3>{safe_html(card['name'])} <span class="priority priority-{priority_class}">{safe_html(card.get('priority', ''))}</span></h3>
            <div class="meta">
                {safe_html(card.get('tier', ''))}
                | Overall Score {safe_html(card.get('score', ''))}
                | Strategic Score {safe_html(card.get('strategic_score', ''))}
                | Enrollment {safe_html(card.get('enrollment', ''))}
            </div>
            <div class="lead"><strong>Lead with:</strong> {safe_html(card.get('lead', ''))}</div>
            <div>{badges}</div>
        </div>
    """, unsafe_allow_html=True)

    render_quick_prep(card)

    if conference_mode:
        col_a, col_b = st.columns([1, 1])
        with col_a:
            st.markdown("**Top 3 Signals**")
            for item in card.get("signals", [])[:3]:
                st.markdown(f"- {item}")
            st.markdown("**Key Contacts**")
            for item in card.get("contacts", [])[:4]:
                st.markdown(f"- {item}")
        with col_b:
            st.markdown("**Top Questions**")
            for item in card.get("questions", [])[:3]:
                st.markdown(f"- **{item}**")
            st.markdown("**Avoid**")
            for item in card.get("avoid", [])[:3]:
                st.markdown(f"- {item}")
    else:
        with st.expander("Top Strategic Signals", expanded=True):
            for item in card.get("signals", []):
                st.markdown(f"- {item}")
        with st.expander("Best-Fit PCG / Emerald Alignment", expanded=False):
            for item in card.get("alignment", []):
                st.markdown(f"- {item}")
        with st.expander("Key Contacts", expanded=False):
            contacts = card.get("contacts", [])
            if contacts:
                for item in contacts:
                    st.markdown(f"- {item}")
            else:
                st.markdown("_No contacts were available for this district._")
        with st.expander("Ask These NEPQ-Style Questions", expanded=True):
            for item in card.get("questions", []):
                st.markdown(f"- **{item}**")
        with st.expander("Listen For / Avoid", expanded=False):
            st.markdown("**Listen for**")
            visible = card.get("listen", [])[:16]
            hidden = card.get("listen", [])[16:]
            st.markdown("".join(chip_html(item) for item in visible), unsafe_allow_html=True)
            if hidden:
                with st.expander("Show all listen-for cues", expanded=False):
                    st.markdown("".join(chip_html(item) for item in hidden), unsafe_allow_html=True)
            st.markdown("**Avoid**")
            for item in card.get("avoid", []):
                st.markdown(f"- {item}")

    quick_prep = (
        f"{card['name']} Quick Prep\n\n"
        f"Lead with: {card.get('lead', '')}\n\n"
        f"Best opening question: {(card.get('questions') or [''])[0]}\n\n"
        f"Listen for: {', '.join(card.get('listen', [])[:10])}\n\n"
        f"Avoid: {(card.get('avoid') or [''])[0]}"
    )
    full_prep = (
        f"{card['name']} Conversation Prep\n\n"
        f"Lead with: {card.get('lead', '')}\n\n"
        "Ask:\n- " + "\n- ".join(card.get("questions", []))
        + "\n\nListen for: " + ", ".join(card.get("listen", []))
        + "\n\nAvoid:\n- " + "\n- ".join(card.get("avoid", []))
    )
    with st.expander("Copy Quick Prep", expanded=False):
        st.text_area("Quick prep copy", quick_prep, height=140, key=f"quick_{card['name']}")
    with st.expander("Copy Full Prep", expanded=False):
        st.text_area("Full prep copy", full_prep, height=190, key=f"full_{card['name']}")

    st.download_button(
        "Download this district brief",
        data=build_docx([card]),
        file_name=f"{card['name'].replace(' ', '_')}_Brief.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"download_{card['name']}",
    )


# ============================================================
# WORD EXPORT
# ============================================================

def build_docx(cards):
    doc = Document()
    title = doc.add_heading("Strategic District Field Guide", 0)
    for run in title.runs:
        run.font.color.rgb = RGBColor(23, 54, 93)
    doc.add_paragraph("Mobile conversation cards for conference prep: lead-with angle, strategic signals, contacts, PCG/Emerald alignment, and NEPQ-style questions.")
    for card in cards:
        doc.add_heading(card["name"], level=1)
        doc.add_paragraph(f"{card.get('tier', '')} | Overall Score {card.get('score', '')} | Strategic Score {card.get('strategic_score', '')} | Enrollment {card.get('enrollment', '')} | Priority {card.get('priority', '')}")
        doc.add_heading("Lead With", level=2)
        doc.add_paragraph(card.get("lead", ""))
        for section_title, key in [("Top Strategic Signals", "signals"), ("Best-Fit PCG / Emerald Alignment", "alignment"), ("Key Contacts", "contacts"), ("NEPQ-Style Questions", "questions"), ("Listen For", "listen"), ("Avoid", "avoid")]:
            doc.add_heading(section_title, level=2)
            for item in card.get(key, []):
                doc.add_paragraph(item, style="List Bullet")
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ============================================================
# HOW-TO GUIDE
# ============================================================

def render_how_to():
    st.markdown('<div class="section-title">How-To Guide</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="howto-box">
    <strong>What this app does</strong><br>
    The Strategic District Field Guide turns the Texas Top 24 Research workbook into searchable district conversation cards for conference prep, quick account refresh, and district leader meetings.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="howto-box">
    <strong>60-second conference workflow</strong>
    <ol>
      <li>Upload the larger Texas Top 24 Research workbook.</li>
      <li>Search the district name.</li>
      <li>Read the Quick Prep box.</li>
      <li>Open Key Contacts if needed.</li>
      <li>Use the Best Opening Question to start the conversation.</li>
      <li>Listen for the cue words and avoid product-first positioning.</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Useful searches")
    st.code("Dallas\nFort Worth\nSPED\nMTSS\nEureka\nBluebonnet\ncareer\nteacher burden", language="text")

    st.markdown("### Tag legend")
    legend = {
        "Math": "Math growth, numeracy, STAAR math, Algebra readiness, student practice.",
        "MTSS": "Intervention, Tier 2/Tier 3, progress monitoring, campus variation.",
        "SPED/ELL": "Special education, multilingual learners, subgroup access, differentiated supports.",
        "CCMR": "College, career, and military readiness; CTE; pathways; career awareness.",
        "Teacher Capacity": "Professional learning, coaching, PLCs, teacher burden, implementation support.",
        "Curriculum / HQIM": "Curriculum adoption, HQIM, Bluebonnet, Eureka, implementation fidelity.",
        "Funding / Grants": "Title funds, grants, LASSO, implementation funding alignment.",
        "Existing Relationship": "A relationship or contract signal exists in the workbook.",
    }
    for tag, meaning in legend.items():
        st.markdown(f"{badge_html(tag)} {meaning}", unsafe_allow_html=True)

    st.markdown("### Export tips")
    st.markdown("- **Download shown cards as Word guide** exports the currently filtered card set.")
    st.markdown("- **Download this district brief** exports a one-district Word brief from the selected card.")
    st.markdown("- The app reads the workbook at runtime; district data does not need to live in GitHub.")


# ============================================================
# OPPORTUNITY MATRIX
# ============================================================

def render_opportunity_matrix(cards):
    st.markdown('<div class="section-title">Opportunity Matrix</div>', unsafe_allow_html=True)
    st.markdown('<div class="helper-text">Use this view to compare districts quickly before choosing which cards to open.</div>', unsafe_allow_html=True)
    if not cards:
        st.info("No cards to show.")
        return
    df = pd.DataFrame([
        {
            "District": c.get("name", ""),
            "Tier": c.get("tier", ""),
            "Overall Score": c.get("score", ""),
            "Strategic Score": c.get("strategic_score", ""),
            "Enrollment": c.get("enrollment", ""),
            "Priority": c.get("priority", ""),
            "Tags": ", ".join(c.get("tags", [])),
            "Lead With": c.get("lead", ""),
        }
        for c in cards
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download matrix as CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="district_opportunity_matrix.csv",
        mime="text/csv",
    )


# ============================================================
# APP LAYOUT
# ============================================================

st.markdown('<div class="main-title">Strategic District Field Guide</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Mobile-friendly district conversation cards for conference prep, strategic account refresh, and quick meeting readiness.</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("1. Data")
    uploaded_file = st.file_uploader("Upload the larger Texas Top 24 Research workbook", type=["xlsx"])
    st.caption("Expected larger workbook includes Master Scorecard, Strategic Indicators, Basic District Info, District Contacts, Leadership and Governance, CSI, and TSI.")

if not uploaded_file:
    tab_field, tab_howto, tab_matrix, tab_diag = st.tabs(["Field Guide", "How-To Guide", "Opportunity Matrix", "Workbook Diagnostics"])
    with tab_field:
        st.info("Upload the larger Texas Top 24 Research workbook to generate district conversation cards.")
    with tab_howto:
        render_how_to()
    with tab_matrix:
        st.info("Upload a workbook to populate the opportunity matrix.")
    with tab_diag:
        st.info("Upload a workbook to view diagnostics.")
    st.stop()

(cards, xls, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, audit_df, score_cols) = load_cards_from_workbook(uploaded_file)

if not cards:
    tab_field, tab_howto, tab_matrix, tab_diag = st.tabs(["Field Guide", "How-To Guide", "Opportunity Matrix", "Workbook Diagnostics"])
    with tab_field:
        st.warning("No eligible district cards were generated. Review the Workbook Diagnostics tab.")
    with tab_howto:
        render_how_to()
    with tab_matrix:
        st.info("No eligible cards available.")
    with tab_diag:
        show_workbook_debug(xls, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, audit_df, score_cols)
    st.stop()

all_districts = [card.get("name", "") for card in cards]
all_tiers = sorted({card.get("tier", "") for card in cards if card.get("tier")})
all_tags = sorted({tag for card in cards for tag in card.get("tags", [])})

with st.sidebar:
    st.header("2. Search")
    query = st.text_input("Search", placeholder="District, contact, signal, offering...")
    st.header("3. Filters")
    selected_tiers = st.multiselect("Tier", options=all_tiers, default=[])
    selected_tags = st.multiselect("Strategic / solution tags", options=all_tags, default=[])
    st.header("4. Shortlist")
    shortlist = st.multiselect("Shortlist districts", options=all_districts, default=[])
    show_shortlist_only = st.checkbox("Show shortlisted only", value=False)
    st.header("5. Display")
    conference_mode = st.toggle("Conference Mode", value=False, help="Shows a shorter card for quick phone-based prep.")
    st.caption("Tip: On a phone, search the district name, then skim Quick Prep, Key Contacts, and Best Opening Question.")

filtered_cards = filter_cards(cards, query, selected_tiers, selected_tags, shortlist, show_shortlist_only)

tab_field, tab_howto, tab_matrix, tab_diag = st.tabs(["Field Guide", "How-To Guide", "Opportunity Matrix", "Workbook Diagnostics"])

with tab_field:
    metric1, metric2, metric3, metric4 = st.columns(4)
    metric1.metric("📍 Cards Shown", len(filtered_cards))
    metric2.metric("✅ Eligible", len(cards))
    metric3.metric("🔥 High Priority", sum(1 for card in filtered_cards if card.get("priority") in ["High", "Very High"]))
    metric4.metric("⭐ Tier 1", sum(1 for card in filtered_cards if card.get("tier") == "Tier 1"))

    if query:
        st.markdown(f'<div class="metric-note">Showing results for: <strong>{safe_html(query)}</strong></div>', unsafe_allow_html=True)

    col_dl1, col_dl2 = st.columns([1, 1])
    with col_dl1:
        st.download_button(
            "Download shown cards as Word guide",
            data=build_docx(filtered_cards),
            file_name="Strategic_District_Field_Guide.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            disabled=not filtered_cards,
        )
    with col_dl2:
        selected_brief = st.selectbox("Download one-district brief", options=[""] + [c["name"] for c in filtered_cards])
        if selected_brief:
            selected_card = next(c for c in filtered_cards if c["name"] == selected_brief)
            st.download_button(
                "Download selected district brief",
                data=build_docx([selected_card]),
                file_name=f"{selected_brief.replace(' ', '_')}_Brief.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_selected_brief",
            )

    st.divider()

    if not filtered_cards:
        st.info("No district cards match the current search/filter.")
    else:
        for card in filtered_cards:
            render_card(card, conference_mode=conference_mode)
            st.divider()

with tab_howto:
    render_how_to()

with tab_matrix:
    render_opportunity_matrix(filtered_cards)

with tab_diag:
    show_workbook_debug(xls, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, audit_df, score_cols)
