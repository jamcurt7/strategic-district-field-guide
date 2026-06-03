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
    --navy:#17365d;
    --blue:#1f4e79;
    --light:#eaf3fb;
    --border:#dbe3ef;
}
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1120px; }
.main-title { font-size: 1.9rem; font-weight: 800; color: var(--navy); margin-bottom: .25rem; }
.subtitle { color: #64748b; font-size: .95rem; margin-bottom: 1rem; }
.card { border:1px solid var(--border); border-radius:18px; padding:1rem; margin-bottom:1rem; background:#ffffff; box-shadow:0 3px 12px rgba(23,54,93,.07); }
.card h3 { color:var(--navy); margin-top:0; margin-bottom:.2rem; }
.meta { color:#64748b; font-size:.85rem; margin-bottom:.65rem; }
.lead { background:var(--light); border-left:4px solid var(--blue); border-radius:12px; padding:.75rem; margin:.6rem 0; }
.badge { display:inline-block; border-radius:999px; padding:.2rem .5rem; margin:.12rem; background:#eef2ff; color:#3730a3; font-size:.75rem; font-weight:700; }
.priority { display:inline-block; border-radius:999px; padding:.25rem .55rem; font-size:.75rem; font-weight:800; }
.priority-very-high { background:#fee2e2; color:#b91c1c; }
.priority-high { background:#dcfce7; color:#166534; }
.priority-medium-high { background:#fef3c7; color:#b45309; }
.priority-medium { background:#e0f2fe; color:#075985; }
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
        base = normalize_text(col)

        if not base:
            base = "Unnamed"

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

    df.columns = make_unique_columns(
        [
            str(c).strip().replace("\n", " ").replace("\r", " ")
            for c in df.columns
        ]
    )

    drop_cols = [
        c
        for c in df.columns
        if normalize_key(c).startswith("methods:")
        or normalize_key(c).startswith("unnamed")
    ]

    if drop_cols:
        df = df.drop(columns=drop_cols)

    return df


def find_sheet_name(xls, desired_name):
    lookup = {normalize_key(sheet): sheet for sheet in xls.sheet_names}
    return lookup.get(normalize_key(desired_name))


def read_table_sheet(xls, desired_sheet_name):
    """
    Reads a sheet by finding the row that contains 'District Name'.

    This handles the larger workbook's Strategic Indicators sheet, where a
    Methods note may appear near the table header.
    """
    actual_sheet = find_sheet_name(xls, desired_sheet_name)

    if not actual_sheet:
        return pd.DataFrame()

    raw = pd.read_excel(
        xls,
        sheet_name=actual_sheet,
        header=None,
        engine="openpyxl",
    )

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

    normalized_map = {
        normalize_key(col): col
        for col in df.columns
    }

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


# ============================================================
# COLUMN DETECTION
# ============================================================

def detect_scorecard_columns(scorecard_df):
    return {
        "district": find_col(
            scorecard_df,
            exact_names=["District Name"],
        ),
        "enrollment": find_col(
            scorecard_df,
            exact_names=["Enrollment"],
        ),
        "strategic_score": find_col(
            scorecard_df,
            exact_names=["Strategic Weighted Score"],
            contains_all=["strategic", "weighted", "score"],
        ),
        "overall_score": find_col(
            scorecard_df,
            exact_names=[
                "Overall Weighted Score (Strategic+Contract+Relationship)",
                "Overall Weighted Score (Strategic+Relationship Weighted)",
                "Overall Weighted Score",
            ],
            contains_all=["overall", "weighted", "score"],
        ),
        "tier": find_col(
            scorecard_df,
            exact_names=["Tier"],
        ),
        "existing_contracts": find_col(
            scorecard_df,
            exact_names=["Existing Contracts in Region (Yes/No)"],
        ),
        "existing_relationships": find_col(
            scorecard_df,
            exact_names=["Existing Relationships"],
        ),
    }


# ============================================================
# LOOKUPS AND ELIGIBILITY
# ============================================================

def is_substantive_strategic_row(row):
    fields_to_check = [
        "Strategic Plan Themes",
        "Math Improvement Mentioned",
        "Math Priority Strength",
        "Intervention Focus",
        "Intervention Focus Details",
        "Teacher Capacity/PD Focus",
        "Teacher Capacity Details",
        "Career Readiness Mentioned",
        "Career Readiness Details",
        "SPED/ELL Improvement Mentioned",
        "SPED/ELL Details",
        "MTSS/Tiered Support Mentioned",
        "MTSS Details",
        "Curriculum Review/Adoption Activity",
        "Curriculum Details",
        "Active Grants (Yes/No)",
        "Grants Details",
        "Sources",
        "Notes",
    ]

    populated = 0

    for field in fields_to_check:
        if normalize_text(get_value(row, field)):
            populated += 1

    return populated >= 2


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
        csi_counts = (
            csi_df.groupby(
                csi_df[DISTRICT_COL].astype(str).str.upper().str.strip()
            )
            .size()
            .to_dict()
        )

    if not tsi_df.empty and DISTRICT_COL in tsi_df.columns:
        tsi_counts = (
            tsi_df.groupby(
                tsi_df[DISTRICT_COL].astype(str).str.upper().str.strip()
            )
            .size()
            .to_dict()
        )

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
        "SUPERINTENDENT",
        "CHIEF ACADEMIC OFFICER",
        "ASSISTANT SUPERINTENDENT",
        "CURRICULUM DIRECTOR",
        "MATH COORDINATOR",
        "CTE COLLEGE CAREER READINESS DIRECTOR",
        "SPECIAL EDUCATION DIRECTOR",
        "ESL ELL COORDINATOR",
        "DIRECTOR ASSESSMENT DATA",
        "PROFESSIONAL LEARNING DIRECTOR",
        "ELA LITERACY COORDINATOR",
    ]

    df = contacts_df.copy()

    if "Position" in df.columns:
        df["_priority"] = (
            df["Position"]
            .astype(str)
            .str.upper()
            .apply(lambda x: role_priority.index(x) if x in role_priority else 99)
        )

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

    combined_text = " ".join(
        normalize_text(v)
        for v in strategic_row.to_dict().values()
    ).lower()

    if normalize_text(get_value(strategic_row, "Math Improvement Mentioned")) or normalize_text(get_value(strategic_row, "Math Priority Strength")):
        tags.append("Math")

    if normalize_text(get_value(strategic_row, "Intervention Focus")) or normalize_text(get_value(strategic_row, "MTSS/Tiered Support Mentioned")):
        tags.append("MTSS")

    if (
        "sped" in combined_text
        or "special education" in combined_text
        or "english learner" in combined_text
        or "emergent bilingual" in combined_text
        or "ell" in combined_text
        or "multilingual" in combined_text
    ):
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
        alignment.append(
            "Elevation Station Math Games — K–8 math practice, fluency, reinforcement, and engagement."
        )
        alignment.append(
            "Elevation intervention curriculum — targeted K–5 support where Tier 2/Tier 3 math or foundational gaps are present."
        )

    if "MTSS" in tags:
        alignment.append(
            "PCG MTSS consulting — system design, campus implementation, intervention fidelity, and data-to-action routines."
        )

    if "SPED/ELL" in tags:
        alignment.append(
            "PCG SPED / multilingual learner support — inclusive practice, service delivery, access to grade-level instruction, and subgroup progress monitoring."
        )

    if "CCMR" in tags:
        alignment.append(
            "RISE Career & Math Mini Lessons — grades 6–9 career-connected math, pathway awareness, and applied readiness."
        )

    if "Curriculum / HQIM" in tags:
        alignment.append(
            "PCG curriculum/HQIM implementation support — adoption fidelity, coaching, PLC routines, and change management."
        )

    if "Teacher Capacity" in tags:
        alignment.append(
            "PCG professional learning and instructional implementation support — teacher capacity, coaching, data use, and scalable instructional routines."
        )

    if "Funding / Grants" in tags:
        alignment.append(
            "Funding alignment support — connect implementation supports to existing grant, Title, or strategic funding streams where appropriate."
        )

    if not alignment:
        alignment.append(
            "Discovery needed — validate strategic needs, stakeholder priorities, and fit before positioning specific resources."
        )

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

    top_tags = [
        tag
        for tag in tags
        if tag not in ["Existing Relationship", "Funding / Grants"]
    ]

    top_tags_text = ", ".join(top_tags[:4]).lower() or "strategic implementation support"

    return (
        f"Lead with {top_tags_text} {relationship_phrase}. "
        f"Frame the conversation around helping {district_name} move from stated priorities "
        "to consistent campus-level execution without adding unnecessary burden."
    )


def build_questions(card):
    tags = card.get("tags", [])

    questions = []

    if "Math" in tags:
        questions.append(
            "How are math goals being translated into weekly instructional decisions, student practice routines, and progress monitoring?"
        )
    else:
        questions.append(
            "How are the district’s strategic priorities being translated into consistent campus-level routines?"
        )

    if "MTSS" in tags:
        questions.append(
            "When students are identified for additional support, where does the process tend to slow down — grouping, scheduling, intervention materials, progress monitoring, or teacher capacity?"
        )
    else:
        questions.append(
            "Where do you see the biggest gap between the strategic plan and day-to-day classroom execution?"
        )

    if "SPED/ELL" in tags:
        questions.append(
            "Where do students with disabilities or emergent bilingual students most often lose access to grade-level expectations?"
        )

    if "Teacher Capacity" in tags or "Curriculum / HQIM" in tags:
        questions.append(
            "After initial training or rollout, where do teachers tend to need the most help — planning, pacing, differentiation, student practice, or responding to data?"
        )

    if "CCMR" in tags:
        questions.append(
            "How early are students connecting academic skills, especially math, to future pathways and readiness expectations?"
        )

    questions.append(
        "If current implementation barriers remain, what are the implications for students, staff, accountability outcomes, and community confidence?"
    )

    questions.append(
        "What would make an external partner feel like implementation support rather than another initiative?"
    )

    questions.append(
        "If you were to pilot targeted support, which campuses, grade bands, or student groups would create the clearest proof point?"
    )

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
    avoid = [
        "Do not lead with a product pitch.",
        "Avoid positioning support as a replacement for the district’s current strategy or adopted curriculum.",
    ]

    if "Curriculum / HQIM" in tags:
        avoid.append(
            "Avoid implying the district needs another curriculum; frame support around implementation, practice, and adoption fidelity."
        )

    if "SPED/ELL" in tags:
        avoid.append(
            "Avoid treating subgroup performance as a side issue; connect it to access, instruction, service delivery, and progress monitoring."
        )

    if "Existing Relationship" in tags:
        avoid.append(
            "Avoid ignoring the existing relationship context; build from credibility and continuity."
        )

    return avoid


# ============================================================
# CARD BUILDING
# ============================================================

def build_card(
    strategic_row,
    score_row,
    basic_lookup,
    contacts_lookup,
    leadership_lookup,
    csi_counts,
    tsi_counts,
    score_cols,
):
    district_name = normalize_text(get_value(strategic_row, DISTRICT_COL))
    key = district_key(district_name)

    raw_score = score_row.get(score_cols["overall_score"], "")
    score = format_number(raw_score, 2)

    strategic_score = (
        format_number(score_row.get(score_cols["strategic_score"], ""), 2)
        if score_cols.get("strategic_score")
        else ""
    )

    tier = normalize_text(score_row.get(score_cols["tier"], ""))

    enrollment = (
        format_enrollment(score_row.get(score_cols["enrollment"], ""))
        if score_cols.get("enrollment")
        else ""
    )

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
        ("schools", "Number of Schools"),
        ("grade span", "Grade Span Served"),
        ("setting", "Urban/Suburban/Rural"),
        ("economically disadvantaged", "% Economically Disadvantaged"),
        ("English learner", "% English Learner"),
        ("special education", "% Special Education"),
        ("major student group", "Major Student Groups"),
        ("student growth trend", "Student Growth Trend"),
    ]

    for label, col in context_fields:
        value = normalize_text(get_value(basic, col))

        if value:
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
        signals.append(
            f"Relationship context: existing contracts in region = {existing_contracts}; existing relationships = {existing_relationships}."
        )
    elif existing_relationships:
        signals.append(
            f"Relationship context: existing relationships = {existing_relationships}."
        )
    elif existing_contracts:
        signals.append(
            f"Relationship context: existing contracts in region = {existing_contracts}."
        )

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
        return (
            [],
            xls,
            strategic_df,
            scorecard_df,
            basic_df,
            contacts_df,
            leadership_df,
            csi_df,
            tsi_df,
            pd.DataFrame(),
            {},
        )

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

        tier = (
            normalize_text(score_row.get(score_cols.get("tier"), ""))
            if matched and score_cols.get("tier")
            else ""
        )

        overall = (
            score_row.get(score_cols.get("overall_score"), "")
            if matched and score_cols.get("overall_score")
            else ""
        )

        eligible = substantive and matched and bool(tier) and bool(normalize_text(overall))

        audit_rows.append(
            {
                "District": district_name,
                "Substantive Strategic Indicators": substantive,
                "Matched Master Scorecard": matched,
                "Detected Tier Column": score_cols.get("tier"),
                "Tier": tier,
                "Detected Overall Score Column": score_cols.get("overall_score"),
                "Overall Score": normalize_text(overall),
                "Eligible": eligible,
            }
        )

        if eligible:
            cards.append(
                build_card(
                    strategic_row,
                    score_row,
                    basic_lookup,
                    contacts_lookup,
                    leadership_lookup,
                    csi_counts,
                    tsi_counts,
                    score_cols,
                )
            )

    audit_df = pd.DataFrame(audit_rows)

    return (
        cards,
        xls,
        strategic_df,
        scorecard_df,
        basic_df,
        contacts_df,
        leadership_df,
        csi_df,
        tsi_df,
        audit_df,
        score_cols,
    )


# ============================================================
# DEBUG
# ============================================================

def show_workbook_debug(
    xls,
    strategic_df,
    scorecard_df,
    basic_df,
    contacts_df,
    leadership_df,
    csi_df,
    tsi_df,
    audit_df,
    score_cols,
):
    with st.expander("Workbook debug details", expanded=False):
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
            st.markdown(f"### {label}")
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
    values = [
        card.get("name", ""),
        card.get("tier", ""),
        str(card.get("score", "")),
        str(card.get("strategic_score", "")),
        card.get("enrollment", ""),
        card.get("priority", ""),
        card.get("lead", ""),
    ]

    for key in ["tags", "signals", "alignment", "contacts", "questions", "listen", "avoid"]:
        values.extend(card.get(key, []))

    return " ".join(map(str, values)).lower()


def filter_cards(cards, query, selected_tiers, selected_tags):
    query = query.strip().lower()
    filtered = []

    for card in cards:
        if selected_tiers and card.get("tier") not in selected_tiers:
            continue

        if selected_tags and not any(tag in card.get("tags", []) for tag in selected_tags):
            continue

        if query and query not in search_blob(card):
            continue

        filtered.append(card)

    return filtered


# ============================================================
# RENDERING
# ============================================================

def render_card(card):
    priority_class = card.get("priority", "Medium").lower().replace(" ", "-")

    badges = "".join(
        [
            f'<span class="badge">{safe_html(tag)}</span>'
            for tag in card.get("tags", [])
        ]
    )

    st.markdown(
        f"""
        <div class="card">
            <h3>{safe_html(card['name'])}</h3>
            <div class="meta">
                {safe_html(card.get('tier', ''))}
                | Overall Score {safe_html(card.get('score', ''))}
                | Strategic Score {safe_html(card.get('strategic_score', ''))}
                | Enrollment {safe_html(card.get('enrollment', ''))}
                &nbsp;<span class="priority priority-{priority_class}">{safe_html(card.get('priority', ''))}</span>
            </div>
            <div class="lead"><strong>Lead with:</strong> {safe_html(card.get('lead', ''))}</div>
            <div>{badges}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

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

        for item in card.get("listen", []):
            st.markdown(f"- {item}")

        st.markdown("**Avoid**")

        for item in card.get("avoid", []):
            st.markdown(f"- {item}")

    prep_text = (
        f"{card['name']} Conversation Prep\n\n"
        f"Lead with: {card.get('lead', '')}\n\n"
        "Ask:\n- "
        + "\n- ".join(card.get("questions", []))
        + "\n\nListen for: "
        + ", ".join(card.get("listen", []))
        + "\n\nAvoid:\n- "
        + "\n- ".join(card.get("avoid", []))
    )

    st.text_area(
        "Copy-ready prep",
        prep_text,
        height=170,
        key=f"prep_{card['name']}",
    )


# ============================================================
# WORD EXPORT
# ============================================================

def build_docx(cards):
    doc = Document()

    title = doc.add_heading("Strategic District Field Guide", 0)

    for run in title.runs:
        run.font.color.rgb = RGBColor(23, 54, 93)

    doc.add_paragraph(
        "Mobile conversation cards for conference prep: lead-with angle, strategic signals, contacts, PCG/Emerald alignment, and NEPQ-style questions."
    )

    for card in cards:
        doc.add_heading(card["name"], level=1)

        doc.add_paragraph(
            f"{card.get('tier', '')} | Overall Score {card.get('score', '')} | Strategic Score {card.get('strategic_score', '')} | Enrollment {card.get('enrollment', '')} | Priority {card.get('priority', '')}"
        )

        doc.add_heading("Lead With", level=2)
        doc.add_paragraph(card.get("lead", ""))

        for section_title, key in [
            ("Top Strategic Signals", "signals"),
            ("Best-Fit PCG / Emerald Alignment", "alignment"),
            ("Key Contacts", "contacts"),
            ("NEPQ-Style Questions", "questions"),
            ("Listen For", "listen"),
            ("Avoid", "avoid"),
        ]:
            doc.add_heading(section_title, level=2)

            for item in card.get(key, []):
                doc.add_paragraph(item, style="List Bullet")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)

    return bio


# ============================================================
# APP LAYOUT
# ============================================================

st.markdown(
    '<div class="main-title">Strategic District Field Guide</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">Searchable mobile-friendly conversation cards for conference prep, district leader meetings, and quick account refresh.</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("1. Upload Workbook")

    uploaded_file = st.file_uploader(
        "Upload the larger Texas Top 24 Research workbook",
        type=["xlsx"],
    )

    st.caption(
        "Expected larger workbook includes Master Scorecard, Strategic Indicators, Basic District Info, District Contacts, Leadership and Governance, CSI, and TSI."
    )

if not uploaded_file:
    st.info("Upload the larger Texas Top 24 Research workbook to generate district conversation cards.")
    st.stop()


(
    cards,
    xls,
    strategic_df,
    scorecard_df,
    basic_df,
    contacts_df,
    leadership_df,
    csi_df,
    tsi_df,
    audit_df,
    score_cols,
) = load_cards_from_workbook(uploaded_file)

show_workbook_debug(
    xls,
    strategic_df,
    scorecard_df,
    basic_df,
    contacts_df,
    leadership_df,
    csi_df,
    tsi_df,
    audit_df,
    score_cols,
)

if not cards:
    st.warning(
        "No eligible district cards were generated. Review the Eligibility audit inside Workbook debug details."
    )

    st.stop()


all_tiers = sorted({card.get("tier", "") for card in cards if card.get("tier")})
all_tags = sorted({tag for card in cards for tag in card.get("tags", [])})

with st.sidebar:
    st.header("2. Search & Filter")

    query = st.text_input(
        "Search",
        placeholder="District, contact, signal, offering...",
    )

    selected_tiers = st.multiselect(
        "Tier",
        options=all_tiers,
        default=[],
    )

    selected_tags = st.multiselect(
        "Strategic / solution tags",
        options=all_tags,
        default=[],
    )

    st.divider()

    st.caption(
        "Tip: On a phone, search the district name, then skim Lead With, Ask These, and Listen For."
    )


filtered_cards = filter_cards(
    cards,
    query,
    selected_tiers,
    selected_tags,
)


col1, col2, col3, col4 = st.columns(4)

col1.metric("Cards shown", len(filtered_cards))
col2.metric("Total eligible cards", len(cards))
col3.metric(
    "High priority shown",
    sum(
        1
        for card in filtered_cards
        if card.get("priority") in ["High", "Very High"]
    ),
)
col4.metric(
    "Tier 1 shown",
    sum(1 for card in filtered_cards if card.get("tier") == "Tier 1"),
)


st.download_button(
    "Download shown cards as Word guide",
    data=build_docx(filtered_cards),
    file_name="Strategic_District_Field_Guide.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    disabled=not filtered_cards,
)

st.divider()


if not filtered_cards:
    st.info("No district cards match the current search/filter.")
else:
    for card in filtered_cards:
        render_card(card)
        st.divider()
