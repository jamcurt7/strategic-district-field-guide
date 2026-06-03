from io import BytesIO
import html

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
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------
# EXACT WORKBOOK STRUCTURE
# ------------------------------------------------------------

SHEET_STRATEGIC = "Strategic Indicators"
SHEET_SCORECARD = "Master Scorecard"
SHEET_BASIC = "Basic District Info"
SHEET_LEADERSHIP = "Leadership and Governance"
SHEET_CONTACTS = "District Contacts"

COL_DISTRICT = "District Name"

# Current Master Scorecard score column from your Streamlit debug output
COL_OVERALL_SCORE = "Overall Weighted Score (Strategic+Relationship Weighted)"

# Fallback old score column, in case an older workbook is uploaded later
COL_OVERALL_SCORE_OLD = "Overall Weighted Score (Strategic+Contract+Relationship)"


# ------------------------------------------------------------
# BASIC HELPERS
# ------------------------------------------------------------

def safe_html(value):
    return html.escape(str(value))


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key(value):
    return (
        normalize_text(value)
        .lower()
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("  ", " ")
        .strip()
    )


def district_key(value):
    return normalize_text(value).upper().strip()


def clean_columns(df):
    """
    Cleans column names and drops the extra Methods column that appears
    in Strategic Indicators.
    """
    if df.empty:
        return df

    df = df.copy()

    df.columns = [
        str(col).strip().replace("\n", " ").replace("\r", " ")
        for col in df.columns
    ]

    methods_cols = [
        col for col in df.columns
        if str(col).strip().lower().startswith("methods:")
    ]

    if methods_cols:
        df = df.drop(columns=methods_cols)

    return df


def find_sheet_name(xls, desired_name):
    lookup = {normalize_key(sheet): sheet for sheet in xls.sheet_names}
    return lookup.get(normalize_key(desired_name))


def read_table_sheet(xls, desired_sheet_name):
    """
    Reads a worksheet even when the true header row is not the first row.

    This matters because Strategic Indicators has a Methods row before
    the real table.
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
        values = [normalize_text(v) for v in row.tolist()]
        if COL_DISTRICT in values:
            header_row_index = idx
            break

    if header_row_index is None:
        return pd.DataFrame()

    headers = raw.iloc[header_row_index].tolist()
    data = raw.iloc[header_row_index + 1:].copy()
    data.columns = headers
    data = clean_columns(data)

    data = data.dropna(how="all")

    if COL_DISTRICT in data.columns:
        data = data[data[COL_DISTRICT].notna()]
        data = data[data[COL_DISTRICT].astype(str).str.strip() != ""]

    return data.reset_index(drop=True)


def format_number(value, decimals=2):
    if pd.isna(value):
        return ""

    try:
        return f"{float(value):.{decimals}f}"
    except Exception:
        return normalize_text(value)


def format_enrollment(value):
    if pd.isna(value):
        return ""

    try:
        return f"{int(float(value)):,}"
    except Exception:
        return normalize_text(value)


def get_score_value(score_row):
    """
    Pulls the overall score from your current workbook column.
    Falls back to the older workbook column if needed.
    """
    if score_row is None:
        return ""

    current = score_row.get(COL_OVERALL_SCORE, "")

    if not pd.isna(current) and normalize_text(current) != "":
        return current

    old = score_row.get(COL_OVERALL_SCORE_OLD, "")

    if not pd.isna(old) and normalize_text(old) != "":
        return old

    return ""


def get_row_value(row, column_name):
    """
    Safely gets a value from a row.
    """
    if row is None:
        return ""

    if column_name in row.index:
        return row.get(column_name, "")

    normalized_lookup = {
        normalize_key(col): col for col in row.index
    }

    lookup_key = normalize_key(column_name)

    if lookup_key in normalized_lookup:
        return row.get(normalized_lookup[lookup_key], "")

    return ""


# ------------------------------------------------------------
# DEBUG PANEL
# ------------------------------------------------------------

def show_workbook_debug(
    xls,
    strategic_df,
    scorecard_df,
    basic_df,
    leadership_df,
    contacts_df,
    eligibility_df,
):
    with st.expander("Workbook debug details", expanded=False):
        st.markdown("### Sheets found")
        st.write(xls.sheet_names)

        debug_items = [
            ("Strategic Indicators", strategic_df),
            ("Master Scorecard", scorecard_df),
            ("Basic District Info", basic_df),
            ("Leadership and Governance", leadership_df),
            ("District Contacts", contacts_df),
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

        st.markdown("### Eligibility audit")
        if eligibility_df.empty:
            st.warning("Eligibility audit is empty. That means no Strategic Indicators rows were processed.")
        else:
            st.dataframe(eligibility_df, use_container_width=True)


# ------------------------------------------------------------
# ELIGIBILITY
# ------------------------------------------------------------

def is_substantive_strategic_row(row):
    """
    A row is substantive if it has actual strategy text in Strategic Indicators.

    This excludes later rows where only District Name exists.
    """
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
        "MTSS/Tiered Support Mentioned",
        "MTSS Details",
        "Curriculum Review/Adoption Activity",
        "Curriculum Details",
        "Active Grants (Yes/No)",
        "Grants Details",
        "Notes",
    ]

    populated_count = 0

    for field in fields_to_check:
        if normalize_text(get_row_value(row, field)):
            populated_count += 1

    return populated_count >= 2


def build_score_lookup(scorecard_df):
    lookup = {}

    if scorecard_df.empty or COL_DISTRICT not in scorecard_df.columns:
        return lookup

    for _, row in scorecard_df.iterrows():
        district_name = normalize_text(get_row_value(row, COL_DISTRICT))

        if district_name:
            lookup[district_key(district_name)] = row

    return lookup


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


# ------------------------------------------------------------
# BASIC DISTRICT CONTEXT
# ------------------------------------------------------------

def get_basic_context(basic_df, district_name):
    if basic_df.empty or COL_DISTRICT not in basic_df.columns:
        return {}

    subset = basic_df[
        basic_df[COL_DISTRICT].astype(str).str.strip().str.upper()
        == district_key(district_name)
    ]

    if subset.empty:
        return {}

    return subset.iloc[0].to_dict()


# ------------------------------------------------------------
# CONTACTS
# ------------------------------------------------------------

def build_contacts_from_district_contacts(contacts_df, district_name, max_contacts=6):
    if contacts_df.empty or COL_DISTRICT not in contacts_df.columns:
        return []

    subset = contacts_df[
        contacts_df[COL_DISTRICT].astype(str).str.strip().str.upper()
        == district_key(district_name)
    ].copy()

    if subset.empty:
        return []

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
    if leadership_df.empty or COL_DISTRICT not in leadership_df.columns:
        return []

    subset = leadership_df[
        leadership_df[COL_DISTRICT].astype(str).str.strip().str.upper()
        == district_key(district_name)
    ]

    if subset.empty:
        return []

    row = subset.iloc[0]

    contacts = []

    superintendent = normalize_text(get_row_value(row, "Superintendent"))
    curriculum_lead = normalize_text(get_row_value(row, "Curriculum Lead"))
    cte_lead = normalize_text(get_row_value(row, "CTE Lead"))
    math_lead = normalize_text(get_row_value(row, "Math Lead"))

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
        contact = contact.replace("(mailto:", " | ")
        contact = contact.replace(")", "")
        cleaned.append(contact)

    return cleaned[:max_contacts]


def build_contacts(contacts_df, leadership_df, district_name):
    contacts = build_contacts_from_district_contacts(contacts_df, district_name)

    if contacts:
        return contacts

    return build_contacts_from_leadership(leadership_df, district_name)


# ------------------------------------------------------------
# TAGS, ALIGNMENT, QUESTIONS
# ------------------------------------------------------------

def infer_tags(strategic_row, score_row):
    tags = []

    all_text = " ".join(
        normalize_text(strategic_row.get(col, ""))
        for col in strategic_row.index
    ).lower()

    if normalize_text(get_row_value(strategic_row, "Math Improvement Mentioned")) or normalize_text(get_row_value(strategic_row, "Math Priority Strength")):
        tags.append("Math")

    if normalize_text(get_row_value(strategic_row, "Intervention Focus")) or normalize_text(get_row_value(strategic_row, "MTSS/Tiered Support Mentioned")):
        tags.append("MTSS")

    if (
        "sped" in all_text
        or "special education" in all_text
        or "english learner" in all_text
        or "emergent bilingual" in all_text
        or "ell" in all_text
        or "multilingual" in all_text
    ):
        tags.append("SPED/ELL")

    if normalize_text(get_row_value(strategic_row, "Career Readiness Mentioned")) or normalize_text(get_row_value(strategic_row, "Career Readiness Details")):
        tags.append("CCMR")

    if normalize_text(get_row_value(strategic_row, "Teacher Capacity/PD Focus")) or normalize_text(get_row_value(strategic_row, "Teacher Capacity Details")):
        tags.append("Teacher Capacity")

    if normalize_text(get_row_value(strategic_row, "Curriculum Review/Adoption Activity")) or normalize_text(get_row_value(strategic_row, "Curriculum Details")):
        tags.append("Curriculum / HQIM")

    if normalize_text(get_row_value(strategic_row, "Active Grants (Yes/No)")) or normalize_text(get_row_value(strategic_row, "Grants Details")):
        tags.append("Funding / Grants")

    existing_relationship = normalize_text(get_row_value(score_row, "Existing Relationships"))

    if existing_relationship.lower() == "yes":
        tags.append("Existing Relationship")

    if not tags:
        tags.append("Strategic Review")

    # Remove duplicates while preserving order
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
    avoid = [
        "Do not lead with a product pitch.",
        "Avoid positioning support as a replacement for the district’s current strategy or adopted curriculum.",
    ]

    if "Curriculum / HQIM" in tags:
        avoid.append("Avoid implying the district needs another curriculum; frame support around implementation, practice, and adoption fidelity.")

    if "SPED/ELL" in tags:
        avoid.append("Avoid treating subgroup performance as a side issue; connect it to access, instruction, service delivery, and progress monitoring.")

    if "Existing Relationship" in tags:
        avoid.append("Avoid ignoring the existing relationship context; build from credibility and continuity.")

    return avoid


# ------------------------------------------------------------
# CARD BUILDER
# ------------------------------------------------------------

def build_card(strategic_row, score_row, basic_df, leadership_df, contacts_df):
    district_name = normalize_text(get_row_value(strategic_row, COL_DISTRICT))

    raw_score = get_score_value(score_row)
    score = format_number(raw_score, decimals=2)
    strategic_score = format_number(get_row_value(score_row, "Strategic Weighted Score"), decimals=2)
    tier = normalize_text(get_row_value(score_row, "Tier"))
    enrollment = format_enrollment(get_row_value(score_row, "Enrollment"))

    tags = infer_tags(strategic_row, score_row)
    priority = determine_priority(tier, raw_score)

    basic = get_basic_context(basic_df, district_name)

    signals = []

    themes = normalize_text(get_row_value(strategic_row, "Strategic Plan Themes"))
    math_strength = normalize_text(get_row_value(strategic_row, "Math Priority Strength"))
    intervention_details = normalize_text(get_row_value(strategic_row, "Intervention Focus Details"))
    teacher_details = normalize_text(get_row_value(strategic_row, "Teacher Capacity Details"))
    career_details = normalize_text(get_row_value(strategic_row, "Career Readiness Details"))
    mtss_details = normalize_text(get_row_value(strategic_row, "MTSS Details"))
    curriculum_details = normalize_text(get_row_value(strategic_row, "Curriculum Details"))
    grants_details = normalize_text(get_row_value(strategic_row, "Grants Details"))

    if themes:
        signals.append(f"Strategic themes: {themes}")

    if math_strength:
        signals.append(f"Math signal: {math_strength}")

    context_parts = []

    for label, col in [
        ("schools", "Number of Schools"),
        ("grade span", "Grade Span Served"),
        ("setting", "Urban/Suburban/Rural"),
        ("economically disadvantaged", "% Economically Disadvantaged"),
        ("English learner", "% English Learner"),
        ("special education", "% Special Education"),
        ("major student group", "Major Student Groups"),
        ("student growth trend", "Student Growth Trend"),
    ]:
        value = normalize_text(basic.get(col, ""))
        if value:
            context_parts.append(f"{label}: {value}")

    if context_parts:
        signals.append("District context: " + "; ".join(context_parts) + ".")

    if intervention_details:
        signals.append(f"Intervention signal: {intervention_details}")

    if mtss_details:
        signals.append(f"MTSS signal: {mtss_details}")

    if teacher_details:
        signals.append(f"Teacher capacity signal: {teacher_details}")

    if career_details:
        signals.append(f"CCMR / career readiness signal: {career_details}")

    if curriculum_details:
        signals.append(f"Curriculum / implementation signal: {curriculum_details}")

    if grants_details:
        signals.append(f"Funding / grants signal: {grants_details}")

    existing_relationships = normalize_text(get_row_value(score_row, "Existing Relationships"))

    if existing_relationships:
        signals.append(f"Relationship context: existing relationships = {existing_relationships}.")

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
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

    strategic_df = read_table_sheet(xls, SHEET_STRATEGIC)
    scorecard_df = read_table_sheet(xls, SHEET_SCORECARD)
    basic_df = read_table_sheet(xls, SHEET_BASIC)
    leadership_df = read_table_sheet(xls, SHEET_LEADERSHIP)
    contacts_df = read_table_sheet(xls, SHEET_CONTACTS)

    if strategic_df.empty or scorecard_df.empty:
        return [], xls, strategic_df, scorecard_df, basic_df, leadership_df, contacts_df, pd.DataFrame()

    score_lookup = build_score_lookup(scorecard_df)

    eligibility_rows = []
    cards = []

    for _, strategic_row in strategic_df.iterrows():
        district_name = normalize_text(get_row_value(strategic_row, COL_DISTRICT))

        if not district_name:
            continue

        substantive = is_substantive_strategic_row(strategic_row)
        score_row = score_lookup.get(district_key(district_name))

        matched_scorecard = score_row is not None
        tier = normalize_text(get_row_value(score_row, "Tier")) if matched_scorecard else ""
        overall_score = get_score_value(score_row) if matched_scorecard else ""

        tier_present = bool(tier)
        score_present = normalize_text(overall_score) != ""

        eligible = substantive and matched_scorecard and tier_present and score_present

        eligibility_rows.append(
            {
                "District": district_name,
                "Substantive Strategic Indicators": substantive,
                "Matched Master Scorecard": matched_scorecard,
                "Tier": tier,
                "Overall Score": normalize_text(overall_score),
                "Eligible": eligible,
            }
        )

        if eligible:
            cards.append(
                build_card(
                    strategic_row=strategic_row,
                    score_row=score_row,
                    basic_df=basic_df,
                    leadership_df=leadership_df,
                    contacts_df=contacts_df,
                )
            )

    eligibility_df = pd.DataFrame(eligibility_rows)

    return cards, xls, strategic_df, scorecard_df, basic_df, leadership_df, contacts_df, eligibility_df


# ------------------------------------------------------------
# SEARCH AND FILTER
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# RENDERING
# ------------------------------------------------------------

def render_card(card):
    priority_class = card.get("priority", "Medium").lower().replace(" ", "-")

    badges = "".join(
        [f'<span class="badge">{safe_html(tag)}</span>' for tag in card.get("tags", [])]
    )

    st.markdown(
        f"""
        <div class="card">
            <h3>{safe_html(card["name"])}</h3>
            <div class="meta">
                {safe_html(card.get("tier", ""))}
                | Overall Score {safe_html(card.get("score", ""))}
                | Strategic Score {safe_html(card.get("strategic_score", ""))}
                | Enrollment {safe_html(card.get("enrollment", ""))}
                &nbsp;
                <span class="priority priority-{priority_class}">
                    {safe_html(card.get("priority", ""))}
                </span>
            </div>
            <div class="lead"><strong>Lead with:</strong> {safe_html(card.get("lead", ""))}</div>
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
            st.markdown("_No District Contacts sheet was found, and no Leadership and Governance contacts were available for this district._")

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
    strategic_df,
    scorecard_df,
    basic_df,
    leadership_df,
    contacts_df,
    eligibility_df,
) = load_cards_from_workbook(uploaded_file)

show_workbook_debug(
    xls=xls,
    strategic_df=strategic_df,
    scorecard_df=scorecard_df,
    basic_df=basic_df,
    leadership_df=leadership_df,
    contacts_df=contacts_df,
    eligibility_df=eligibility_df,
)

if not cards:
    st.warning(
        "No eligible district cards were generated. Review the Eligibility audit inside Workbook debug details."
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
