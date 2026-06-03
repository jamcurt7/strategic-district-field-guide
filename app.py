from io import BytesIO
from pathlib import Path

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
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------

def normalize_text(value):
    """
    Safely converts spreadsheet values to clean strings.
    """
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_col(df, candidates):
    """
    Finds a column in a dataframe even if spacing/case differs.
    """
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in normalized:
            return normalized[key]
    return None


def get_sheet(xls, sheet_name):
    """
    Reads a worksheet if it exists.
    Returns an empty dataframe if it does not.
    """
    if sheet_name in xls.sheet_names:
        return pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
    return pd.DataFrame()


def has_substantive_indicator(row):
    """
    Determines whether a district has enough Strategic Indicators data
    to include in the field guide.
    """
    fields = [
        "Strategic Plan Themes",
        "Math Improvement Mentioned ",
        "Math Priority Strength",
        "Intervention Focus",
        "Intervention Focus Details",
        "Teacher Capacity/PD Focus ",
        "Teacher Capacity Details",
        "Career Readiness Mentioned ",
        "Career Readiness Details",
        "SPED/ELL Improvement Mentioned",
        "SPED/ELL Details",
        "MTSS/Tiered Support Mentioned",
        "MTSS Details",
        "Curriculum Review/Adoption Activity ",
        "Curriculum Details",
    ]

    for field in fields:
        if normalize_text(row.get(field, "")):
            return True

    return False


def value_contains(value, terms):
    """
    Checks whether a spreadsheet field contains any of the target terms.
    """
    text = normalize_text(value).lower()
    return any(term.lower() in text for term in terms)


def build_contacts(contacts_df, district_name, max_contacts=6):
    """
    Pulls high-value contacts for a district.
    Prioritizes superintendent, academic, math, CTE, SPED, ELL,
    accountability, and professional learning roles.
    """
    if contacts_df.empty:
        return []

    district_col = find_col(contacts_df, ["District Name"])
    if not district_col:
        return []

    subset = contacts_df[
        contacts_df[district_col].astype(str).str.strip().str.upper()
        == district_name.strip().upper()
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
    ]

    position_col = find_col(subset, ["Position"])

    if position_col:
        subset["_priority"] = subset[position_col].astype(str).str.upper().apply(
            lambda x: role_priority.index(x) if x in role_priority else 99
        )
        subset = subset.sort_values("_priority")

    contacts = []

    for _, r in subset.head(max_contacts).iterrows():
        first = normalize_text(r.get("First Name", ""))
        last = normalize_text(r.get("Last Name", ""))
        title = normalize_text(r.get("Title", r.get("Position", "")))

        name = f"{first} {last}".strip()

        if name and title:
            contacts.append(f"{name} — {title}")
        elif title:
            contacts.append(title)

    return contacts


def extract_basic_context(basic_df, district_name):
    """
    Pulls basic district context such as CSI/TSI counts and demographics.
    """
    if basic_df.empty:
        return {}

    district_col = find_col(basic_df, ["District Name"])
    if not district_col:
        return {}

    subset = basic_df[
        basic_df[district_col].astype(str).str.strip().str.upper()
        == district_name.strip().upper()
    ]

    if subset.empty:
        return {}

    return subset.iloc[0].to_dict()


def infer_tags(indicator_row, score_row):
    """
    Creates filter tags based on the district's strategic indicators.
    """
    tags = []

    if value_contains(indicator_row.get("Math Priority Strength", ""), ["high", "strong", "very", "extreme"]):
        tags.append("Math")

    if value_contains(indicator_row.get("Intervention Focus", ""), ["yes"]):
        tags.append("MTSS")

    if value_contains(indicator_row.get("SPED/ELL Improvement Mentioned", ""), ["yes", "priority", "high", "extreme"]):
        tags.append("SPED/ELL")

    if value_contains(indicator_row.get("Career Readiness Mentioned ", ""), ["yes"]):
        tags.append("CCMR")

    if value_contains(indicator_row.get("Teacher Capacity/PD Focus ", ""), ["yes"]):
        tags.append("Teacher Capacity")

    if value_contains(indicator_row.get("Curriculum Review/Adoption Activity ", ""), ["yes", "adoption", "hqim", "bluebonnet", "eureka"]):
        tags.append("HQIM")

    if normalize_text(score_row.get("Existing Relationships", "")).lower() == "yes":
        tags.append("Existing Relationship")

    if not tags:
        tags.append("Strategic Review")

    return tags


def determine_priority(tier, score):
    """
    Converts score/tier context into a simple conference-use priority label.
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


def build_alignment(tags):
    """
    Maps strategic tags to PCG / Emerald alignment language.
    """
    alignment = []

    if "Math" in tags:
        alignment.append(
            "Elevation Station — K–8 math practice, fluency, engagement, and reinforcement"
        )

    if "MTSS" in tags:
        alignment.append(
            "PCG MTSS — system design, campus implementation, intervention fidelity, and data-to-action routines"
        )

    if "SPED/ELL" in tags:
        alignment.append(
            "PCG SPED / multilingual learner support — access, compliance-to-instruction alignment, and subgroup supports"
        )

    if "CCMR" in tags:
        alignment.append(
            "RISE Career & Math Mini Lessons — grade 6–9 career-connected math and pathway awareness"
        )

    if "HQIM" in tags or "Teacher Capacity" in tags:
        alignment.append(
            "PCG implementation and professional learning — coaching, PLC routines, curriculum adoption fidelity, and change management"
        )

    if not alignment:
        alignment.append(
            "Discovery needed — validate strategic fit and stakeholder priorities"
        )

    return alignment


def build_questions(tags):
    """
    Generates NEPQ-style questions based on district tags.
    """
    questions = []

    # Situation questions
    if "Math" in tags:
        questions.append(
            "How are math goals being translated into weekly instructional decisions and progress monitoring routines?"
        )
    else:
        questions.append(
            "How are your current strategic priorities being translated into consistent campus-level routines?"
        )

    # Problem awareness
    questions.append(
        "Where do you see the biggest gap between the district plan and day-to-day classroom execution?"
    )

    if "MTSS" in tags:
        questions.append(
            "When students are identified for support, where does the process tend to slow down — grouping, scheduling, materials, progress monitoring, or teacher capacity?"
        )

    if "SPED/ELL" in tags:
        questions.append(
            "Where do students with disabilities or emergent bilingual students most often lose access to grade-level expectations?"
        )

    if "HQIM" in tags or "Teacher Capacity" in tags:
        questions.append(
            "After initial curriculum or instructional training, where do teachers tend to need the most help — planning, pacing, differentiation, or responding to data?"
        )

    # Impact / consequence
    questions.append(
        "If current implementation barriers remain, what are the implications for students, staff, accountability outcomes, and community confidence?"
    )

    # Strategic vision
    questions.append(
        "What would make an external partner feel like implementation support rather than another initiative?"
    )

    return questions


def build_lead_with(tags):
    """
    Creates a short lead-with angle for a conference conversation.
    """
    tag_phrase = " / ".join(tags[:4]).lower()

    return (
        f"Lead with consultative support around {tag_phrase}, "
        "emphasizing implementation coherence, practical classroom use, and reduced teacher burden."
    )


def build_card(indicator_row, score_row, contacts_df, basic_df):
    """
    Combines data from workbook sheets into one district conversation card.
    """
    district_name = normalize_text(indicator_row.get("District Name", ""))

    tier = normalize_text(score_row.get("Tier", "")) or "Unscored"

    score = score_row.get("Overall Weighted Score (Strategic+Contract+Relationship)", "")

    try:
        score = round(float(score), 2)
    except Exception:
        score = ""

    enrollment = score_row.get("Enrollment", "")

    if pd.notna(enrollment):
        try:
            enrollment = f"{int(float(enrollment)):,}"
        except Exception:
            enrollment = str(enrollment)
    else:
        enrollment = ""

    basic_context = extract_basic_context(basic_df, district_name)

    csi = normalize_text(basic_context.get("CSI Schools", ""))
    tsi = normalize_text(basic_context.get("TSI Schools", ""))

    csi_tsi_line = (
        f"CSI/TSI context: {csi} CSI and {tsi} TSI campuses."
        if csi or tsi
        else "Accountability context should be validated from CSI/TSI sheets."
    )

    themes = normalize_text(indicator_row.get("Strategic Plan Themes", ""))
    math_strength = normalize_text(indicator_row.get("Math Priority Strength", ""))
    intervention_details = normalize_text(indicator_row.get("Intervention Focus Details", ""))
    sped_ell_details = normalize_text(indicator_row.get("SPED/ELL Details", ""))
    teacher_details = normalize_text(indicator_row.get("Teacher Capacity Details", ""))
    career_details = normalize_text(indicator_row.get("Career Readiness Details", ""))
    curriculum_details = normalize_text(indicator_row.get("Curriculum Details", ""))

    tags = infer_tags(indicator_row, score_row)
    priority = determine_priority(tier, score)

    signals = []

    if themes:
        signals.append(themes)
    else:
        signals.append("Strategic indicators suggest a need for additional qualitative review.")

    if math_strength:
        signals.append(f"Math signal: {math_strength}")

    signals.append(csi_tsi_line)

    if intervention_details:
        signals.append(f"Intervention signal: {intervention_details[:280]}")

    if sped_ell_details:
        signals.append(f"SPED/ELL signal: {sped_ell_details[:280]}")

    if teacher_details:
        signals.append(f"Teacher capacity signal: {teacher_details[:260]}")

    if career_details and "CCMR" in tags:
        signals.append(f"CCMR signal: {career_details[:260]}")

    if curriculum_details and "HQIM" in tags:
        signals.append(f"Curriculum / adoption signal: {curriculum_details[:260]}")

    card = {
        "name": district_name,
        "tier": tier,
        "score": score,
        "enrollment": enrollment,
        "priority": priority,
        "tags": tags,
        "lead": build_lead_with(tags),
        "signals": signals,
        "alignment": build_alignment(tags),
        "contacts": build_contacts(contacts_df, district_name),
        "questions": build_questions(tags),
        "listen": tags + [
            "implementation fidelity",
            "teacher capacity",
            "campus variation",
            "progress monitoring",
        ],
        "avoid": [
            "Do not lead with a product pitch.",
            "Avoid positioning support as a replacement for district strategy or adopted curriculum.",
        ],
    }

    return card


def load_cards_from_workbook(uploaded_file):
    """
    Reads the workbook and builds district cards.
    """
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

    indicators_df = get_sheet(xls, "Strategic Indicators")
    scorecard_df = get_sheet(xls, "Master Scorecard")
    contacts_df = get_sheet(xls, "District Contacts")
    basic_df = get_sheet(xls, "Basic District Info")

    if indicators_df.empty or scorecard_df.empty:
        st.error("Workbook must include Strategic Indicators and Master Scorecard sheets.")
        return []

    indicators_district_col = find_col(indicators_df, ["District Name"])
    scorecard_district_col = find_col(scorecard_df, ["District Name"])

    if not indicators_district_col or not scorecard_district_col:
        st.error("Could not find District Name columns in Strategic Indicators or Master Scorecard.")
        return []

    cards = []

    for _, indicator_row in indicators_df.iterrows():
        district_name = normalize_text(indicator_row.get(indicators_district_col, ""))

        if not district_name:
            continue

        if not has_substantive_indicator(indicator_row):
            continue

        score_matches = scorecard_df[
            scorecard_df[scorecard_district_col].astype(str).str.strip().str.upper()
            == district_name.upper()
        ]

        if score_matches.empty:
            continue

        score_row = score_matches.iloc[0].to_dict()

        if not normalize_text(score_row.get("Tier", "")):
            continue

        cards.append(build_card(indicator_row, score_row, contacts_df, basic_df))

    return cards


def search_blob(card):
    """
    Combines all searchable card text into one lowercase string.
    """
    values = [
        card.get("name", ""),
        card.get("tier", ""),
        str(card.get("score", "")),
        card.get("enrollment", ""),
        card.get("priority", ""),
        card.get("lead", ""),
    ]

    for key in ["tags", "signals", "alignment", "contacts", "questions", "listen", "avoid"]:
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


def render_card(card):
    """
    Renders one district card in Streamlit.
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
                {card.get("tier", "")} | Score {card.get("score", "")} | Enrollment {card.get("enrollment", "")}
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
            st.markdown("_No contacts found for this district in the workbook._")

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
    )

    st.text_area(
        "Copy-ready prep",
        prep_text,
        height=150,
        key=f"prep_{card['name']}",
    )


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
            f"{card.get('tier', '')} | Score {card.get('score', '')} | Enrollment {card.get('enrollment', '')} | Priority {card.get('priority', '')}"
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
        "Expected sheets: Strategic Indicators, Master Scorecard, District Contacts, Basic District Info."
    )


if not uploaded_file:
    st.info("Upload the Excel workbook to generate district conversation cards.")
    st.stop()


cards = load_cards_from_workbook(uploaded_file)

if not cards:
    st.warning("No eligible district cards were generated. Check that Strategic Indicators and Master Scorecard are populated.")
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

col1, col2, col3 = st.columns(3)

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
