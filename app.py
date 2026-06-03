from io import BytesIO

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor


# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------

st.set_page_config(
    page_title="Strategic District Field Guide",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ------------------------------------------------------------
# CUSTOM STYLING
# ------------------------------------------------------------

CUSTOM_CSS = """
<style>
:root {
    --navy:#17365d;
    --blue:#1f4e79;
    --light:#eaf3fb;
    --border:#dbe3ef;
}

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1100px;
}

.main-title {
    font-size: 1.9rem;
    font-weight: 800;
    color: var(--navy);
    margin-bottom: .25rem;
}

.subtitle {
    color: #64748b;
    font-size: .95rem;
    margin-bottom: 1rem;
}

.card {
    border:1px solid var(--border);
    border-radius:18px;
    padding:1rem;
    margin-bottom:1rem;
    background:#ffffff;
    box-shadow:0 3px 12px rgba(23,54,93,.07);
}

.card h3 {
    color:var(--navy);
    margin-top:0;
    margin-bottom:.2rem;
}

.meta {
    color:#64748b;
    font-size:.85rem;
    margin-bottom:.65rem;
}

.lead {
    background:var(--light);
    border-left:4px solid var(--blue);
    border-radius:12px;
    padding:.75rem;
    margin:.6rem 0;
}

.badge {
    display:inline-block;
    border-radius:999px;
    padding:.2rem .5rem;
    margin:.12rem;
    background:#eef2ff;
    color:#3730a3;
    font-size:.75rem;
    font-weight:700;
}

.priority {
    display:inline-block;
    border-radius:999px;
    padding:.25rem .55rem;
    font-size:.75rem;
    font-weight:800;
}

.priority-very-high {
    background:#fee2e2;
    color:#b91c1c;
}

.priority-high {
    background:#dcfce7;
    color:#166534;
}

.priority-medium-high {
    background:#fef3c7;
    color:#b45309;
}

.priority-medium {
    background:#e0f2fe;
    color:#075985;
}

.small-muted {
    color:#64748b;
    font-size:.82rem;
}

.debug-box {
    background:#f8fafc;
    border:1px solid #dbe3ef;
    border-radius:12px;
    padding:0.75rem;
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------
# WORKBOOK CONFIGURATION
# ------------------------------------------------------------

SHEET_STRATEGIC = "Strategic Indicators"
SHEET_SCORECARD = "Master Scorecard"
SHEET_CONTACTS = "District Contacts"
SHEET_BASIC = "Basic District Info"
SHEET_LEADERSHIP = "Leadership and Governance"

COL_DISTRICT = "District Name"

# Strategic Indicators columns
COL_STRATEGIC_THEMES = "Strategic Plan Themes"
COL_MATH_MENTIONED = "Math Improvement Mentioned"
COL_MATH_STRENGTH = "Math Priority Strength"
COL_INTERVENTION_FOCUS = "Intervention Focus"
COL_INTERVENTION_DETAILS = "Intervention Focus Details"
COL_TEACHER_FOCUS = "Teacher Capacity/PD Focus"
COL_TEACHER_DETAILS = "Teacher Capacity Details"
COL_CAREER_MENTIONED = "Career Readiness Mentioned"
COL_CAREER_DETAILS = "Career Readiness Details"
COL_SPED_ELL_MENTIONED = "SPED/ELL Improvement Mentioned"
COL_SPED_ELL_DETAILS = "SPED/ELL Details"
COL_MTSS_MENTIONED = "MTSS/Tiered Support Mentioned"
COL_MTSS_DETAILS = "MTSS Details"
COL_CURRICULUM_ACTIVITY = "Curriculum Review/Adoption Activity"
COL_CURRICULUM_DETAILS = "Curriculum Details"
COL_GRANTS = "Active Grants (Yes/No)"
COL_GRANTS_DETAILS = "Grants Details"
COL_SOURCES = "Sources"
COL_NOTES = "Notes"

# Master Scorecard columns
COL_ENROLLMENT = "Enrollment"
COL_TIER = "Tier"
COL_STRATEGIC_WEIGHTED_SCORE = "Strategic Weighted Score"

# Current workbook score column
COL_OVERALL_SCORE_CURRENT = "Overall Weighted Score (Strategic+Relationship Weighted)"

# Older workbook score column, supported for backward compatibility
COL_OVERALL_SCORE_OLD = "Overall Weighted Score (Strategic+Contract+Relationship)"

COL_EXISTING_RELATIONSHIPS = "Existing Relationships"
COL_EXISTING_CONTRACTS = "Existing Contracts in Region (Yes/No)"


# ------------------------------------------------------------
# BASIC HELPERS
# ------------------------------------------------------------

def normalize_text(value):
    """
    Safely convert any spreadsheet cell value to a clean string.
    """
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key(value):
    """
    Normalize text for matching sheet names, column names, and district names.
    """
    return normalize_text(value).lower().replace("\n", " ").replace("\r", " ").strip()


def clean_columns(df):
    """
    Clean column names by stripping whitespace and converting line breaks.

    This handles columns like:
    - 'Career Readiness Mentioned '
    - 'Teacher Capacity/PD Focus '
    - columns with hidden line breaks
    """
    if df.empty:
        return df

    df = df.copy()
    df.columns = [
        str(col).strip().replace("\n", " ").replace("\r", " ")
        for col in df.columns
    ]
    return df


def find_sheet_name(xls, desired_name):
    """
    Find a sheet by exact name after trimming/case normalization.
    """
    lookup = {normalize_key(sheet): sheet for sheet in xls.sheet_names}
    return lookup.get(normalize_key(desired_name))


def read_table_sheet(xls, desired_sheet_name):
    """
    Reads a worksheet whose real table may not start on row 1.

    This is critical because Strategic Indicators includes a Methods row
    before the real header row containing District Name.
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

    header_row_index = None

    for idx, row in raw.iterrows():
        row_values = [normalize_text(value) for value in row.tolist()]
        if COL_DISTRICT in row_values:
            header_row_index = idx
            break

    if header_row_index is None:
        return pd.DataFrame()

    headers = raw.iloc[header_row_index].tolist()
    data = raw.iloc[header_row_index + 1:].copy()
    data.columns = headers
    data = clean_columns(data)

    # Drop fully empty rows
    data = data.dropna(how="all")

    # Drop rows where District Name is blank
    if COL_DISTRICT in data.columns:
        data = data[data[COL_DISTRICT].notna()]
        data = data[data[COL_DISTRICT].astype(str).str.strip() != ""]

    return data.reset_index(drop=True)


def value_contains(value, terms):
    """
    Checks whether a field contains any of the target terms.
    """
    text = normalize_text(value).lower()
    return any(term.lower() in text for term in terms)


def format_number(value, decimals=2):
    """
    Format numeric values cleanly.
    """
    if pd.isna(value):
        return ""

    try:
        number = float(value)
        return f"{number:.{decimals}f}"
    except Exception:
        return normalize_text(value)


def format_enrollment(value):
    """
    Format enrollment as 56,400 instead of 56400.
    """
    if pd.isna(value):
        return ""

    try:
        return f"{int(float(value)):,}"
    except Exception:
        return normalize_text(value)


def district_match_key(value):
    """
    Normalize district names for cross-sheet matching.
    """
    return normalize_text(value).upper()


def get_overall_score(score_row):
    """
    Pulls the overall score from either the current workbook column name
    or the older workbook column name.

    Current workbook:
    Overall Weighted Score (Strategic+Relationship Weighted)

    Older workbook:
    Overall Weighted Score (Strategic+Contract+Relationship)
    """
    current_score = score_row.get(COL_OVERALL_SCORE_CURRENT, "")

    if not pd.isna(current_score) and normalize_text(current_score) != "":
        return current_score

    old_score = score_row.get(COL_OVERALL_SCORE_OLD, "")

    if not pd.isna(old_score) and normalize_text(old_score) != "":
        return old_score

    return ""


# ------------------------------------------------------------
# WORKBOOK DEBUG
# ------------------------------------------------------------

def show_workbook_debug(xls, indicators_df, scorecard_df, contacts_df, basic_df, leadership_df):
    """
    Shows what the app sees in the workbook.
    Useful if the workbook changes later.
    """
    with st.expander("Workbook debug details", expanded=False):
        st.markdown("### Sheets found")
        st.write(xls.sheet_names)

        debug_items = [
            ("Strategic Indicators", indicators_df),
            ("Master Scorecard", scorecard_df),
            ("District Contacts", contacts_df),
            ("Basic District Info", basic_df),
            ("Leadership and Governance", leadership_df),
        ]

        for label, df in debug_items:
            st.markdown(f"### {label}")
            st.write(f"Rows: {len(df)}")
            st.write("Columns:")
            st.write(list(df.columns))

            if COL_DISTRICT in df.columns:
                st.write("First district names found:")
                st.write(
                    df[COL_DISTRICT]
                    .dropna()
                    .astype(str)
                    .head(15)
                    .tolist()
                )


# ------------------------------------------------------------
# ELIGIBILITY LOGIC
# ------------------------------------------------------------

def has_substantive_indicator(row):
    """
    Determines whether a Strategic Indicators row has actual usable strategy content.

    In the current workbook, rows from Aldine ISD through Fort Worth ISD
    have filled strategy fields, while later rows mostly contain only district names.
    """
    strategic_fields = [
        COL_STRATEGIC_THEMES,
        COL_MATH_MENTIONED,
        COL_MATH_STRENGTH,
        COL_INTERVENTION_FOCUS,
        COL_INTERVENTION_DETAILS,
        COL_TEACHER_FOCUS,
        COL_TEACHER_DETAILS,
        COL_CAREER_MENTIONED,
        COL_CAREER_DETAILS,
        COL_SPED_ELL_MENTIONED,
        COL_SPED_ELL_DETAILS,
        COL_MTSS_MENTIONED,
        COL_MTSS_DETAILS,
        COL_CURRICULUM_ACTIVITY,
        COL_CURRICULUM_DETAILS,
        COL_GRANTS,
        COL_GRANTS_DETAILS,
        COL_SOURCES,
        COL_NOTES,
    ]

    for field in strategic_fields:
        if normalize_text(row.get(field, "")):
            return True

    return False


def is_fully_scored(score_row):
    """
    Determines whether a district has a complete enough Master Scorecard row.

    A district must have:
    - Tier
    - Overall weighted score
    """
    if not score_row:
        return False

    tier = normalize_text(score_row.get(COL_TIER, ""))
    score = get_overall_score(score_row)

    if not tier:
        return False

    if pd.isna(score) or normalize_text(score) == "":
        return False

    return True


# ------------------------------------------------------------
# CONTACT LOGIC
# ------------------------------------------------------------

def build_contacts_from_district_contacts(contacts_df, district_name, max_contacts=6):
    """
    Pulls high-value contacts from District Contacts if that sheet exists.
    """
    if contacts_df.empty or COL_DISTRICT not in contacts_df.columns:
        return []

    subset = contacts_df[
        contacts_df[COL_DISTRICT].astype(str).str.strip().str.upper()
        == district_match_key(district_name)
    ].copy()

    if subset.empty:
        return []

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

    if "Position" in subset.columns:
        subset["_priority"] = subset["Position"].astype(str).str.upper().apply(
            lambda x: role_priority.index(x) if x in role_priority else 99
        )
        subset = subset.sort_values("_priority")

    contacts = []

    for _, row in subset.head(max_contacts).iterrows():
        first = normalize_text(row.get("First Name", ""))
        last = normalize_text(row.get("Last Name", ""))
        title = normalize_text(row.get("Title", ""))
        position = normalize_text(row.get("Position", ""))

        name = f"{first} {last}".strip()

        if name and title:
            contacts.append(f"{name} — {title}")
        elif name and position:
            contacts.append(f"{name} — {position}")
        elif title:
            contacts.append(title)

    return contacts


def build_contacts_from_leadership(leadership_df, district_name, max_contacts=6):
    """
    Pulls fallback contacts from Leadership and Governance.

    Current workbook has Leadership and Governance but may not have District Contacts.
    """
    if leadership_df.empty or COL_DISTRICT not in leadership_df.columns:
        return []

    subset = leadership_df[
        leadership_df[COL_DISTRICT].astype(str).str.strip().str.upper()
        == district_match_key(district_name)
    ]

    if subset.empty:
        return []

    row = subset.iloc[0].to_dict()

    contacts = []

    possible_pairs = [
        ("Superintendent", "Superintendent"),
        ("Curriculum Lead", "Title"),
        ("CTE Lead", "Title"),
        ("Math Lead", "Title"),
    ]

    # Because Leadership and Governance has repeated generic "Title" columns in Excel,
    # pandas may rename duplicates. We handle simple known fields first.
    superintendent = normalize_text(row.get("Superintendent", ""))
    if superintendent:
        contacts.append(f"{superintendent} — Superintendent")

    curriculum_lead = normalize_text(row.get("Curriculum Lead", ""))
    if curriculum_lead:
        contacts.append(f"{curriculum_lead} — Curriculum / Academic Lead")

    cte_lead = normalize_text(row.get("CTE Lead", ""))
    if cte_lead:
        contacts.append(f"{cte_lead} — CTE / Career Readiness Lead")

    math_lead = normalize_text(row.get("Math Lead", ""))
    if math_lead:
        contacts.append(f"{math_lead} — Math Lead")

    # Clean up accidental mailto/link text if present.
    cleaned = []
    for contact in contacts:
        contact = contact.replace("[", "").replace("]", "")
        contact = contact.replace("(mailto:", " | ")
        contact = contact.replace(")", "")
        cleaned.append(contact)

    return cleaned[:max_contacts]


def build_contacts(contacts_df, leadership_df, district_name):
    """
    Uses District Contacts first, then Leadership and Governance as fallback.
    """
    contacts = build_contacts_from_district_contacts(contacts_df, district_name)

    if contacts:
        return contacts

    return build_contacts_from_leadership(leadership_df, district_name)


# ------------------------------------------------------------
# BASIC DISTRICT INFO LOGIC
# ------------------------------------------------------------

def get_basic_context(basic_df, district_name):
    """
    Pulls one row from Basic District Info.
    """
    if basic_df.empty or COL_DISTRICT not in basic_df.columns:
        return {}

    subset = basic_df[
        basic_df[COL_DISTRICT].astype(str).str.strip().str.upper()
        == district_match_key(district_name)
    ]

    if subset.empty:
        return {}

    return subset.iloc[0].to_dict()


# ------------------------------------------------------------
# STRATEGIC TAGGING
# ------------------------------------------------------------

def infer_tags(indicator_row, score_row):
    """
    Creates searchable filter tags based on Strategic Indicators and Master Scorecard.
    """
    tags = []

    math_mentioned = normalize_text(indicator_row.get(COL_MATH_MENTIONED, ""))
    math_strength = normalize_text(indicator_row.get(COL_MATH_STRENGTH, ""))

    if math_mentioned or math_strength:
        tags.append("Math")

    intervention_focus = normalize_text(indicator_row.get(COL_INTERVENTION_FOCUS, ""))
    mtss_mentioned = normalize_text(indicator_row.get(COL_MTSS_MENTIONED, ""))

    if intervention_focus or mtss_mentioned:
        tags.append("MTSS")

    sped_ell = normalize_text(indicator_row.get(COL_SPED_ELL_MENTIONED, ""))
    sped_ell_details = normalize_text(indicator_row.get(COL_SPED_ELL_DETAILS, ""))

    if sped_ell or sped_ell_details:
        tags.append("SPED/ELL")

    career = normalize_text(indicator_row.get(COL_CAREER_MENTIONED, ""))
    career_details = normalize_text(indicator_row.get(COL_CAREER_DETAILS, ""))

    if career or career_details:
        tags.append("CCMR")

    teacher = normalize_text(indicator_row.get(COL_TEACHER_FOCUS, ""))
    teacher_details = normalize_text(indicator_row.get(COL_TEACHER_DETAILS, ""))

    if teacher or teacher_details:
        tags.append("Teacher Capacity")

    curriculum = normalize_text(indicator_row.get(COL_CURRICULUM_ACTIVITY, ""))
    curriculum_details = normalize_text(indicator_row.get(COL_CURRICULUM_DETAILS, ""))

    if curriculum or curriculum_details:
        tags.append("Curriculum / HQIM")

    grants = normalize_text(indicator_row.get(COL_GRANTS, ""))
    grants_details = normalize_text(indicator_row.get(COL_GRANTS_DETAILS, ""))

    if grants or grants_details:
        tags.append("Funding / Grants")

    existing_relationships = normalize_text(score_row.get(COL_EXISTING_RELATIONSHIPS, ""))

    if existing_relationships.lower() == "yes":
        tags.append("Existing Relationship")

    if not tags:
        tags.append("Strategic Review")

    return tags


def determine_priority(tier, score):
    """
    Converts tier/score into a simple conference-use priority label.
    """
    if tier == "Tier 1":
        return "Very High"

    if tier == "Tier 2":
        return "High"

    try:
        numeric_score = float(score)

        if numeric_score >= 3.5:
            return "Medium-High"
    except Exception:
        pass

    return "Medium"


# ------------------------------------------------------------
# STRATEGIC LANGUAGE GENERATION
# ------------------------------------------------------------

def build_lead_with(card):
    """
    Creates a short lead-with statement using district-specific signals.
    """
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

    top_tags_text = ", ".join(top_tags[:4]).lower()

    if not top_tags_text:
        top_tags_text = "strategic implementation support"

    return (
        f"Lead with {top_tags_text} {relationship_phrase}. "
        f"Frame the conversation around helping {district_name} move from stated priorities "
        "to consistent campus-level execution without adding unnecessary burden."
    )


def build_alignment(tags):
    """
    Maps strategic tags to PCG / Emerald alignment.
    """
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


def build_questions(card):
    """
    Builds NEPQ-style discovery questions based on the district tags and context.
    """
    tags = card.get("tags", [])

    questions = []

    # Situation
    if "Math" in tags:
        questions.append(
            "How are math goals being translated into weekly instructional decisions, student practice routines, and progress monitoring?"
        )
    else:
        questions.append(
            "How are the district’s strategic priorities being translated into consistent campus-level routines?"
        )

    # Problem awareness
    if "MTSS" in tags:
        questions.append(
            "When students are identified for additional support, where does the process tend to slow down — grouping, scheduling, intervention materials, progress monitoring, or teacher capacity?"
        )
    else:
        questions.append(
            "Where do you see the biggest gap between the strategic plan and day-to-day classroom execution?"
        )

    # SPED / ELL
    if "SPED/ELL" in tags:
        questions.append(
            "Where do students with disabilities or emergent bilingual students most often lose access to grade-level expectations?"
        )

    # Teacher capacity / HQIM
    if "Teacher Capacity" in tags or "Curriculum / HQIM" in tags:
        questions.append(
            "After initial training or rollout, where do teachers tend to need the most help — planning, pacing, differentiation, student practice, or responding to data?"
        )

    # CCMR
    if "CCMR" in tags:
        questions.append(
            "How early are students connecting academic skills, especially math, to future pathways and readiness expectations?"
        )

    # Impact / consequence
    questions.append(
        "If current implementation barriers remain, what are the implications for students, staff, accountability outcomes, and community confidence?"
    )

    # Vision
    questions.append(
        "What would make an external partner feel like implementation support rather than another initiative?"
    )

    # Change / implementation
    questions.append(
        "If you were to pilot targeted support, which campuses, grade bands, or student groups would create the clearest proof point?"
    )

    return questions


def build_listen_for(tags):
    """
    Creates listen-for cues for live conference conversations.
    """
    listen_for = []

    tag_map = {
        "Math": [
            "math growth",
            "early numeracy",
            "Algebra readiness",
            "STAAR math",
            "student practice",
        ],
        "MTSS": [
            "intervention fidelity",
            "Tier 2",
            "Tier 3",
            "progress monitoring",
            "campus variation",
        ],
        "SPED/ELL": [
            "access to grade-level instruction",
            "service delivery",
            "emergent bilingual students",
            "students with disabilities",
            "subgroup gaps",
        ],
        "CCMR": [
            "pathways",
            "industry-based certifications",
            "TSIA2",
            "dual credit",
            "career awareness",
        ],
        "Teacher Capacity": [
            "teacher burden",
            "coaching",
            "professional learning",
            "PLC routines",
            "instructional consistency",
        ],
        "Curriculum / HQIM": [
            "adoption fidelity",
            "HQIM",
            "curriculum implementation",
            "Eureka",
            "Bluebonnet",
            "instructional materials",
        ],
        "Funding / Grants": [
            "Title funding",
            "grant alignment",
            "LASSO",
            "federal funds",
            "implementation funding",
        ],
        "Existing Relationship": [
            "existing relationship",
            "current contract",
            "expansion",
            "trusted partner",
        ],
    }

    for tag in tags:
        listen_for.extend(tag_map.get(tag, []))

    listen_for.extend(
        [
            "data cycles",
            "implementation barriers",
            "capacity constraints",
        ]
    )

    unique = []

    for item in listen_for:
        if item not in unique:
            unique.append(item)

    return unique


def build_avoid(tags):
    """
    Creates avoid cues.
    """
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


# ------------------------------------------------------------
# CARD BUILDER
# ------------------------------------------------------------

def build_card(indicator_row, score_row, contacts_df, basic_df, leadership_df):
    """
    Combines the workbook sheets into a single district conversation card.
    """
    district_name = normalize_text(indicator_row.get(COL_DISTRICT, ""))

    tier = normalize_text(score_row.get(COL_TIER, ""))
    raw_score = get_overall_score(score_row)
    score = format_number(raw_score, decimals=2)

    strategic_score = format_number(
        score_row.get(COL_STRATEGIC_WEIGHTED_SCORE, ""),
        decimals=2,
    )

    enrollment = format_enrollment(score_row.get(COL_ENROLLMENT, ""))

    tags = infer_tags(indicator_row, score_row)
    priority = determine_priority(tier, raw_score)

    basic = get_basic_context(basic_df, district_name)

    number_of_schools = normalize_text(basic.get("Number of Schools", ""))
    grade_span = normalize_text(basic.get("Grade Span Served", ""))
    econ = normalize_text(basic.get("% Economically Disadvantaged", ""))
    ell = normalize_text(basic.get("% English Learner", ""))
    sped = normalize_text(basic.get("% Special Education", ""))
    student_groups = normalize_text(basic.get("Major Student Groups", ""))
    district_type = normalize_text(basic.get("Urban/Suburban/Rural", ""))
    growth_trend = normalize_text(basic.get("Student Growth Trend", ""))

    signals = []

    themes = normalize_text(indicator_row.get(COL_STRATEGIC_THEMES, ""))
    math_strength = normalize_text(indicator_row.get(COL_MATH_STRENGTH, ""))
    intervention_details = normalize_text(indicator_row.get(COL_INTERVENTION_DETAILS, ""))
    teacher_details = normalize_text(indicator_row.get(COL_TEACHER_DETAILS, ""))
    career_details = normalize_text(indicator_row.get(COL_CAREER_DETAILS, ""))
    sped_ell_details = normalize_text(indicator_row.get(COL_SPED_ELL_DETAILS, ""))
    mtss_details = normalize_text(indicator_row.get(COL_MTSS_DETAILS, ""))
    curriculum_details = normalize_text(indicator_row.get(COL_CURRICULUM_DETAILS, ""))
    grants_details = normalize_text(indicator_row.get(COL_GRANTS_DETAILS, ""))

    if themes:
        signals.append(f"Strategic themes: {themes}")

    if math_strength:
        signals.append(f"Math signal: {math_strength}")

    context_parts = []

    if district_type:
        context_parts.append(f"{district_type} context")
    if number_of_schools:
        context_parts.append(f"{number_of_schools} schools")
    if grade_span:
        context_parts.append(f"grade span: {grade_span}")
    if student_groups:
        context_parts.append(f"major student group: {student_groups}")
    if econ:
        context_parts.append(f"economically disadvantaged: {econ}")
    if ell:
        context_parts.append(f"English learner: {ell}")
    if sped:
        context_parts.append(f"special education: {sped}")
    if growth_trend:
        context_parts.append(f"student growth trend: {growth_trend}")

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

    existing_contracts = normalize_text(score_row.get(COL_EXISTING_CONTRACTS, ""))
    existing_relationships = normalize_text(score_row.get(COL_EXISTING_RELATIONSHIPS, ""))

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
        "contacts": build_contacts(contacts_df, leadership_df, district_name),
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


# ------------------------------------------------------------
# LOAD WORKBOOK
# ------------------------------------------------------------

def load_cards_from_workbook(uploaded_file):
    """
    Reads the uploaded workbook and generates eligible district cards.
    """
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

    indicators_df = read_table_sheet(xls, SHEET_STRATEGIC)
    scorecard_df = read_table_sheet(xls, SHEET_SCORECARD)
    contacts_df = read_table_sheet(xls, SHEET_CONTACTS)
    basic_df = read_table_sheet(xls, SHEET_BASIC)
    leadership_df = read_table_sheet(xls, SHEET_LEADERSHIP)

    if indicators_df.empty or scorecard_df.empty:
        return [], xls, indicators_df, scorecard_df, contacts_df, basic_df, leadership_df

    if COL_DISTRICT not in indicators_df.columns or COL_DISTRICT not in scorecard_df.columns:
        return [], xls, indicators_df, scorecard_df, contacts_df, basic_df, leadership_df

    score_lookup = {}

    for _, score_row in scorecard_df.iterrows():
        district = normalize_text(score_row.get(COL_DISTRICT, ""))

        if not district:
            continue

        score_lookup[district_match_key(district)] = score_row.to_dict()

    cards = []

    for _, indicator_row in indicators_df.iterrows():
        district_name = normalize_text(indicator_row.get(COL_DISTRICT, ""))

        if not district_name:
            continue

        if not has_substantive_indicator(indicator_row):
            continue

        score_row = score_lookup.get(district_match_key(district_name))

        if not score_row:
            continue

        if not is_fully_scored(score_row):
            continue

        card = build_card(
            indicator_row,
            score_row,
            contacts_df,
            basic_df,
            leadership_df,
        )

        cards.append(card)

    return cards, xls, indicators_df, scorecard_df, contacts_df, basic_df, leadership_df


# ------------------------------------------------------------
# SEARCH AND FILTER
# ------------------------------------------------------------

def search_blob(card):
    """
    Combines all searchable card text into one lowercase string.
    """
    values = [
        card.get("name", ""),
        card.get("tier", ""),
        str(card.get("score", "")),
        str(card.get("strategic_score", "")),
        card.get("enrollment", ""),
        card.get("priority", ""),
        card.get("lead", ""),
    ]

    for key in [
        "tags",
        "signals",
        "alignment",
        "contacts",
        "questions",
        "listen",
        "avoid",
    ]:
        values.extend(card.get(key, []))

    return " ".join(map(str, values)).lower()


def filter_cards(cards, query, selected_tiers, selected_tags):
    """
    Applies search, tier filter, and tag filter.
    """
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


# ------------------------------------------------------------
# RENDERING
# ------------------------------------------------------------

def render_card(card):
    """
    Renders one district conversation card.
    """
    priority_class = card.get("priority", "Medium").lower().replace(" ", "-")

    badges = "".join(
        [f'<span class="badge">{tag}</span>' for tag in card.get("tags", [])]
    )

    st.markdown(
        f"""
        <div class="card">
            <h3>{card["name"]}</h3>
            <div class="meta">
                {card.get("tier", "")}
                | Overall Score {card.get("score", "")}
                | Strategic Score {card.get("strategic_score", "")}
                | Enrollment {card.get("enrollment", "")}
                &nbsp;
                <span class="priority priority-{priority_class}">
                    {card.get("priority", "")}
                </span>
            </div>
            <div class="lead"><strong>Lead with:</strong> {card.get("lead", "")}</div>
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
            st.markdown(
                "_No District Contacts sheet was found, and no Leadership and Governance contacts were available for this district._"
            )

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


# ------------------------------------------------------------
# WORD EXPORT
# ------------------------------------------------------------

def build_docx(cards):
    """
    Builds a Word document from visible cards.
    """
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

        sections = [
            ("Top Strategic Signals", "signals"),
            ("Best-Fit PCG / Emerald Alignment", "alignment"),
            ("Key Contacts", "contacts"),
            ("NEPQ-Style Questions", "questions"),
            ("Listen For", "listen"),
            ("Avoid", "avoid"),
        ]

        for section_title, key in sections:
            doc.add_heading(section_title, level=2)

            for item in card.get(key, []):
                doc.add_paragraph(item, style="List Bullet")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)

    return bio


# ------------------------------------------------------------
# APP LAYOUT
# ------------------------------------------------------------

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
        "Upload district research workbook",
        type=["xlsx"],
    )

    st.caption(
        "Expected workbook: Texas Top 24 Research.xlsx with Strategic Indicators, Master Scorecard, Basic District Info, and Leadership and Governance. District Contacts is optional."
    )

if not uploaded_file:
    st.info("Upload the Excel workbook to generate district conversation cards.")
    st.stop()


(
    cards,
    xls,
    indicators_df,
    scorecard_df,
    contacts_df,
    basic_df,
    leadership_df,
) = load_cards_from_workbook(uploaded_file)

show_workbook_debug(
    xls,
    indicators_df,
    scorecard_df,
    contacts_df,
    basic_df,
    leadership_df,
)

if not cards:
    st.warning(
        "No eligible district cards were generated. The app found the workbook, but no districts met the eligibility rules."
    )

    st.markdown(
        """
        A district must:
        - Appear in **Strategic Indicators**
        - Have substantive Strategic Indicators content
        - Match a district in **Master Scorecard**
        - Have a populated **Tier**
        - Have a populated **Overall Weighted Score**
        """
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
    query=query,
    selected_tiers=selected_tiers,
    selected_tags=selected_tags,
)


# ------------------------------------------------------------
# SUMMARY METRICS
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# EXPORT BUTTON
# ------------------------------------------------------------

st.download_button(
    "Download shown cards as Word guide",
    data=build_docx(filtered_cards),
    file_name="Strategic_District_Field_Guide.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    disabled=not filtered_cards,
)

st.divider()


# ------------------------------------------------------------
# RENDER CARDS
# ------------------------------------------------------------

if not filtered_cards:
    st.info("No district cards match the current search/filter.")
else:
    for card in filtered_cards:
        render_card(card)
        st.divider()
