from io import BytesIO
import html, os, re
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="Strategic District Field Guide", page_icon="📘", layout="wide", initial_sidebar_state="expanded")

BUILT_IN_WORKBOOK = "Texas Top 24 Research.xlsx"
DISTRICT_COL = "District Name"
SHEETS = {
    "strategic": "Strategic Indicators",
    "scorecard": "Master Scorecard",
    "basic": "Basic District Info",
    "contacts": "District Contacts",
    "leadership": "Leadership and Governance",
    "csi": "CSI",
    "tsi": "TSI",
}

CSS = """
<style>
:root{--navy:#102a43;--blue:#1d4ed8;--blue-soft:#eff6ff;--border:#cbd5e1;--text:#0f172a;--muted:#475569;--page:#f3f6fa;--surface:#fff;--slate-soft:#f8fafc;--red:#991b1b;--red-soft:#fef2f2;--green:#166534;--green-soft:#f0fdf4;--amber:#92400e;--amber-soft:#fffbeb}
.stApp,[data-testid="stAppViewContainer"],[data-testid="stHeader"]{background:var(--page)!important;color:var(--text)!important} html,body,p,li,span,div,label,.stMarkdown,.stMarkdown *{color:var(--text)!important}.block-container{padding-top:1.2rem;padding-bottom:2.5rem;max-width:1180px}[data-testid="stSidebar"]{background:#fff!important;border-right:1px solid var(--border)!important}[data-testid="stSidebar"] *{color:var(--text)!important}h1,h2,h3,h4,h5,h6{color:var(--navy)!important}.main-title{font-size:2rem;font-weight:850;color:var(--navy)!important;margin-bottom:.25rem;letter-spacing:-.02em}.subtitle{color:var(--muted)!important;font-size:.98rem;margin-bottom:1rem}.section-title{color:var(--navy)!important;font-size:1.25rem;font-weight:800;margin:.6rem 0 .3rem}.helper-text{color:var(--muted)!important;font-size:.9rem}.stTabs [data-baseweb="tab-list"]{background:#fff!important;border:1px solid var(--border)!important;border-radius:14px!important;padding:.25rem!important}.stTabs [data-baseweb="tab"]{color:#334155!important;border-radius:10px!important;font-weight:700!important}.stTabs [aria-selected="true"]{background:var(--blue-soft)!important;color:var(--blue)!important}[data-testid="stMetric"]{background:#fff!important;border:1px solid var(--border)!important;border-radius:16px!important;padding:.85rem!important;box-shadow:0 4px 14px rgba(15,23,42,.05)!important}[data-testid="stMetric"] *{color:var(--text)!important}input,textarea,select,[data-baseweb="input"],[data-baseweb="select"],[data-baseweb="textarea"]{background:#fff!important;color:var(--text)!important;border-color:var(--border)!important}.stButton button,.stDownloadButton button{background:var(--blue)!important;color:#fff!important;border:1px solid var(--blue)!important;border-radius:10px!important;font-weight:750!important}[data-testid="stExpander"]{background:#fff!important;border:1px solid var(--border)!important;border-radius:14px!important}[data-testid="stExpander"] *{color:var(--text)!important}.card{border:1px solid var(--border);border-radius:20px;padding:1.05rem;margin-bottom:1rem;background:var(--surface);box-shadow:0 8px 24px rgba(15,23,42,.08)}.card h3{color:var(--navy)!important;margin-top:0;margin-bottom:.2rem;font-size:1.35rem}.meta{color:var(--muted)!important;font-size:.88rem;margin-bottom:.75rem;line-height:1.45}.summary-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.75rem;margin:.75rem 0 .6rem}.summary-box{background:#fff;border:1px solid var(--border);border-radius:14px;padding:.75rem;box-shadow:0 3px 12px rgba(15,23,42,.04)}.summary-label{color:var(--muted)!important;font-size:.72rem;font-weight:850;text-transform:uppercase;letter-spacing:.05em;margin-bottom:.2rem}.summary-value{font-size:.92rem;line-height:1.35}.quickprep{background:#fff;border:1px solid #bfdbfe;border-left:5px solid var(--blue);border-radius:16px;padding:.9rem;margin:.7rem 0 .8rem;box-shadow:0 4px 16px rgba(15,23,42,.05)}.quickprep-title{color:var(--navy)!important;font-weight:850;font-size:.95rem;text-transform:uppercase;letter-spacing:.04em;margin-bottom:.25rem}.badge,.chip{display:inline-block;border-radius:999px;padding:.25rem .58rem;margin:.14rem .16rem .14rem 0;font-size:.76rem;font-weight:750;line-height:1.2;border:1px solid transparent}.chip{background:var(--slate-soft);color:#334155!important;border-color:#e2e8f0}.tag-math{background:#dbeafe;color:#1e3a8a!important}.tag-mtss{background:#ccfbf1;color:#134e4a!important}.tag-spedell{background:#ede9fe;color:#4c1d95!important}.tag-ccmr{background:#fef3c7;color:#78350f!important}.tag-teacher{background:#f1f5f9;color:#1e293b!important}.tag-hqim{background:#e0e7ff;color:#312e81!important}.tag-funding{background:#fef9c3;color:#713f12!important}.tag-relationship{background:#dcfce7;color:#14532d!important}.tag-default{background:#eef2ff;color:#312e81!important}.priority{display:inline-block;border-radius:999px;padding:.28rem .62rem;font-size:.75rem;font-weight:850}.priority-very-high{background:var(--red-soft);color:var(--red)!important;border:1px solid #fecaca}.priority-high{background:var(--green-soft);color:var(--green)!important;border:1px solid #bbf7d0}.priority-medium-high{background:var(--amber-soft);color:var(--amber)!important;border:1px solid #fde68a}.priority-medium{background:#e0f2fe;color:#075985!important;border:1px solid #bae6fd}.howto-box{background:#fff;border:1px solid var(--border);border-radius:18px;padding:1rem 1.1rem;margin:.75rem 0;box-shadow:0 4px 16px rgba(15,23,42,.05)}@media(max-width:800px){.summary-grid{grid-template-columns:1fr}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# -----------------------------
# Helpers
# -----------------------------
def safe_html(v): return html.escape(str(v))

def clean(v):
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except Exception:
        pass
    return str(v).strip()

def key(v):
    return re.sub(r"\s+", " ", clean(v).lower().replace("\n", " ").replace("\r", " ").replace("\xa0", " ")).strip()

def dkey(v): return clean(v).upper().strip()

def unique_cols(cols):
    seen, out = {}, []
    for c in cols:
        base = clean(c) or "Unnamed"
        if base not in seen:
            seen[base] = 0; out.append(base)
        else:
            seen[base] += 1; out.append(f"{base}.{seen[base]}")
    return out

def clean_columns(df):
    if df.empty: return df
    df = df.copy(); df.columns = unique_cols([str(c).strip().replace("\n", " ").replace("\r", " ") for c in df.columns])
    drops = [c for c in df.columns if key(c).startswith("methods:") or key(c).startswith("unnamed")]
    return df.drop(columns=drops) if drops else df

def sheet_name(xls, desired):
    m = {key(s): s for s in xls.sheet_names}
    return m.get(key(desired))

def read_sheet(xls, desired):
    actual = sheet_name(xls, desired)
    if not actual: return pd.DataFrame()
    raw = pd.read_excel(xls, sheet_name=actual, header=None, engine="openpyxl")
    if raw.empty: return pd.DataFrame()
    header = None
    for i, row in raw.iterrows():
        if DISTRICT_COL in [clean(v) for v in row.tolist()]:
            header = i; break
    if header is None: return pd.DataFrame()
    data = raw.iloc[header+1:].copy(); data.columns = [clean(v) for v in raw.iloc[header].tolist()]
    data = clean_columns(data).dropna(how="all")
    if DISTRICT_COL in data.columns:
        data = data[data[DISTRICT_COL].notna()]
        data = data[data[DISTRICT_COL].astype(str).str.strip() != ""]
    return data.reset_index(drop=True)

def find_col(df, exact=(), contains=()):
    if df.empty: return None
    norm = {key(c): c for c in df.columns}
    for e in exact:
        if key(e) in norm: return norm[key(e)]
    if contains:
        for c in df.columns:
            ck = key(c)
            if all(t.lower() in ck for t in contains): return c
    return None

def val(row, col, default=""):
    if row is None: return default
    items = row.items() if isinstance(row, dict) else [(k, row.get(k, default)) for k in getattr(row, "index", [])]
    for k, v in items:
        if key(k) == key(col): return v
    return default

def fmt_num(v, d=2):
    if clean(v) == "": return ""
    try: return f"{float(v):.{d}f}"
    except Exception: return clean(v)

def fmt_enroll(v):
    if clean(v) == "": return ""
    try: return f"{int(float(v)):,}"
    except Exception: return clean(v)

def fmt_pct(v, d=0):
    if clean(v) == "": return ""
    try:
        n = float(v); pct = n*100 if abs(n) <= 1 else n
        return f"{pct:.{d}f}%"
    except Exception: return clean(v)

def trunc(text, n=190):
    text = clean(text)
    if len(text) <= n: return text
    return text[:n].rsplit(" ", 1)[0] + "..."

def fmt_contact(name, title="", email="", position=""):
    name, title, email, position = clean(name), clean(title), clean(email), clean(position)
    title = title or position
    if name and title and email: return f"{name} — {title} | {email}"
    if name and title: return f"{name} — {title}"
    if name and email: return f"{name} | {email}"
    if title and email: return f"{title} | {email}"
    return name or title or email

def tag_class(t):
    return {"Math":"tag-math","MTSS":"tag-mtss","SPED/ELL":"tag-spedell","CCMR":"tag-ccmr","Teacher Capacity":"tag-teacher","Curriculum / HQIM":"tag-hqim","Funding / Grants":"tag-funding","Existing Relationship":"tag-relationship"}.get(t,"tag-default")

def badge(t): return f'<span class="badge {tag_class(t)}">{safe_html(t)}</span>'
def chip(t): return f'<span class="chip">{safe_html(t)}</span>'

# -----------------------------
# Data parsing
# -----------------------------
def score_cols(df):
    return {
        "district": find_col(df, exact=[DISTRICT_COL]),
        "enrollment": find_col(df, exact=["Enrollment"]),
        "strategic": find_col(df, exact=["Strategic Weighted Score"], contains=["strategic","weighted","score"]),
        "overall": find_col(df, exact=["Overall Weighted Score (Strategic+Contract+Relationship)", "Overall Weighted Score (Strategic+Relationship Weighted)", "Overall Weighted Score"], contains=["overall","weighted","score"]),
        "tier": find_col(df, exact=["Tier"]),
    }

def substantive(row):
    fields = ["Strategic Plan Themes","Math Improvement Mentioned","Math Priority Strength","Intervention Focus","Intervention Focus Details","Teacher Capacity/PD Focus","Teacher Capacity Details","Career Readiness Mentioned","Career Readiness Details","SPED/ELL Improvement Mentioned","SPED/ELL Details","MTSS/Tiered Support Mentioned","MTSS Details","Curriculum Review/Adoption Activity","Curriculum Details","Active Grants (Yes/No)","Grants Details","Sources","Notes"]
    return sum(1 for f in fields if clean(val(row, f))) >= 2

def lookup_by_district(df):
    out = {}
    if df.empty or DISTRICT_COL not in df.columns: return out
    for _, r in df.iterrows():
        d = clean(r.get(DISTRICT_COL, ""))
        if d: out[dkey(d)] = r.to_dict()
    return out

def counts_by_district(df):
    if df.empty or DISTRICT_COL not in df.columns: return {}
    return df.groupby(df[DISTRICT_COL].astype(str).str.upper().str.strip()).size().to_dict()

def build_contacts_from_contacts(df):
    out = {}
    if df.empty or DISTRICT_COL not in df.columns: return out
    priority = ["SUPERINTENDENT","CHIEF ACADEMIC OFFICER","ASSISTANT SUPERINTENDENT","CURRICULUM DIRECTOR","MATH COORDINATOR","CTE COLLEGE CAREER READINESS DIRECTOR","SPECIAL EDUCATION DIRECTOR","ESL ELL COORDINATOR","DIRECTOR ASSESSMENT DATA","PROFESSIONAL LEARNING DIRECTOR"]
    df = df.copy()
    if "Position" in df.columns:
        df["_p"] = df["Position"].astype(str).str.upper().apply(lambda x: priority.index(x) if x in priority else 99)
        df = df.sort_values([DISTRICT_COL, "_p"])
    for _, r in df.iterrows():
        d = clean(r.get(DISTRICT_COL, ""))
        if not d: continue
        name = f"{clean(r.get('First Name',''))} {clean(r.get('Last Name',''))}".strip()
        c = fmt_contact(name, r.get("Title", ""), r.get("Email", ""), r.get("Position", ""))
        if c:
            out.setdefault(dkey(d), [])
            if c not in out[dkey(d)]: out[dkey(d)].append(c)
    return out

def build_contacts_from_leadership(df):
    out = {}
    if df.empty or DISTRICT_COL not in df.columns: return out
    for _, r in df.iterrows():
        d = clean(r.get(DISTRICT_COL, ""))
        if not d: continue
        contacts = [
            fmt_contact(val(r,"Superintendent"), "Superintendent", val(r,"Email")),
            fmt_contact(val(r,"Curriculum Lead"), val(r,"Title") or "Curriculum / Academic Lead", val(r,"Email.1")),
            fmt_contact(val(r,"CTE Lead"), val(r,"Title.1") or "CTE / Career Readiness Lead", val(r,"Email.2")),
            fmt_contact(val(r,"Math Lead"), val(r,"Title.2") or "Math Lead", val(r,"Email.3")),
        ]
        out[dkey(d)] = [c.replace("[","").replace("]","").replace("(mailto:"," | ").replace(")","") for c in contacts if c][:6]
    return out

def get_contacts(d, contact_lookup, leadership_lookup):
    k = dkey(d)
    return (contact_lookup.get(k) or leadership_lookup.get(k) or [])[:6]

# -----------------------------
# Content builders
# -----------------------------
def infer_tags(strat, score):
    tags, text = [], " ".join(clean(v) for v in strat.to_dict().values()).lower()
    if clean(val(strat,"Math Improvement Mentioned")) or clean(val(strat,"Math Priority Strength")): tags.append("Math")
    if clean(val(strat,"Intervention Focus")) or clean(val(strat,"MTSS/Tiered Support Mentioned")): tags.append("MTSS")
    if any(x in text for x in ["sped","special education","english learner","emergent bilingual","ell","multilingual"]): tags.append("SPED/ELL")
    if clean(val(strat,"Career Readiness Mentioned")) or clean(val(strat,"Career Readiness Details")): tags.append("CCMR")
    if clean(val(strat,"Teacher Capacity/PD Focus")) or clean(val(strat,"Teacher Capacity Details")): tags.append("Teacher Capacity")
    if clean(val(strat,"Curriculum Review/Adoption Activity")) or clean(val(strat,"Curriculum Details")): tags.append("Curriculum / HQIM")
    if clean(val(strat,"Active Grants (Yes/No)")) or clean(val(strat,"Grants Details")): tags.append("Funding / Grants")
    if clean(val(score,"Existing Relationships")).lower() == "yes": tags.append("Existing Relationship")
    if not tags: tags.append("Strategic Review")
    return list(dict.fromkeys(tags))

def alignment(tags):
    a=[]
    if "Math" in tags: a += ["Elevation Station Math Games — K–8 math practice, fluency, reinforcement, and engagement.", "Elevation intervention curriculum — targeted K–5 support where Tier 2/Tier 3 math or foundational gaps are present."]
    if "MTSS" in tags: a.append("PCG MTSS consulting — system design, campus implementation, intervention fidelity, and data-to-action routines.")
    if "SPED/ELL" in tags: a.append("PCG SPED / multilingual learner support — inclusive practice, service delivery, access to grade-level instruction, and subgroup progress monitoring.")
    if "CCMR" in tags: a.append("RISE Career & Math Mini Lessons — grades 6–9 career-connected math, pathway awareness, and applied readiness.")
    if "Curriculum / HQIM" in tags: a.append("PCG curriculum/HQIM implementation support — adoption fidelity, coaching, PLC routines, and change management.")
    if "Teacher Capacity" in tags: a.append("PCG professional learning and instructional implementation support — teacher capacity, coaching, data use, and scalable instructional routines.")
    if "Funding / Grants" in tags: a.append("Funding alignment support — connect implementation supports to existing grant, Title, or strategic funding streams where appropriate.")
    return a or ["Discovery needed — validate strategic needs, stakeholder priorities, and fit before positioning specific resources."]

def priority(tier, score):
    if tier == "Tier 1": return "Very High"
    if tier == "Tier 2": return "High"
    try:
        if float(score) >= 3.5: return "Medium-High"
    except Exception: pass
    return "Medium"

def questions(tags):
    q=[]
    q.append("How are campuses currently using assessment data to decide which students need additional math support?" if "Math" in tags else "How are campuses currently translating district priorities into daily instructional routines?")
    q.append("Where does implementation tend to vary most across campuses, grade levels, or student groups?")
    if "MTSS" in tags: q.append("What tends to slow down the response after students are identified for additional support?")
    if "SPED/ELL" in tags: q.append("Where do students with disabilities or multilingual learners need more consistent access to grade-level expectations?")
    if "Teacher Capacity" in tags or "Curriculum / HQIM" in tags: q.append("What makes it easier or harder for teachers to use new materials or supports consistently after initial training?")
    if "CCMR" in tags: q.append("Where do students begin connecting academic skills to future pathways and career readiness expectations?")
    q += ["What evidence would tell district and campus leaders that a support model is working well enough to expand?", "If a small pilot were considered, what would make the pilot credible to teachers and principals?"]
    return list(dict.fromkeys(q))[:6]

def listen_for(tags):
    m={"Math":["math growth","early numeracy","Algebra readiness","STAAR math","student practice"],"MTSS":["intervention fidelity","Tier 2","Tier 3","progress monitoring","campus variation"],"SPED/ELL":["access to grade-level instruction","service delivery","emergent bilingual students","students with disabilities","subgroup gaps"],"CCMR":["pathways","industry-based certifications","TSIA2","dual credit","career awareness"],"Teacher Capacity":["teacher burden","coaching","professional learning","PLC routines","instructional consistency"],"Curriculum / HQIM":["adoption fidelity","HQIM","curriculum implementation","Eureka","Bluebonnet","instructional materials"],"Funding / Grants":["Title funding","grant alignment","LASSO","federal funds","implementation funding"],"Existing Relationship":["existing relationship","current contract","expansion","trusted partner"]}
    out=[]
    for t in tags: out += m.get(t, [])
    out += ["data cycles","implementation barriers","capacity constraints"]
    return list(dict.fromkeys(out))

def guidance(tags):
    g=["Anchor the conversation to the district’s stated strategic plan, accountability context, and implementation priorities.", "Frame PCG/Emerald support as capacity-building, implementation coherence, and measurable proof-of-concept support."]
    if "Curriculum / HQIM" in tags: g.append("Emphasize implementation fidelity, teacher usability, and alignment to adopted materials rather than a competing curriculum conversation.")
    if "SPED/ELL" in tags: g.append("Connect subgroup performance to instructional access, service delivery, and progress monitoring routines.")
    if "Existing Relationship" in tags: g.append("Reference existing relationship context and look for expansion pathways tied to current district priorities.")
    return g[:4]

def priority_need_opp(tags):
    out=[]
    if "Math" in tags: out.append("Priority: math growth. Need: consistent practice and progress monitoring. Opportunity: Elevation Station and targeted math intervention.")
    if "MTSS" in tags: out.append("Priority: intervention. Need: clear Tier 2/Tier 3 routines. Opportunity: PCG MTSS support and intervention design.")
    if "SPED/ELL" in tags: out.append("Priority: subgroup access. Need: grade-level expectations with usable scaffolds. Opportunity: PCG SPED/ML support and aligned practice tools.")
    if "CCMR" in tags: out.append("Priority: readiness pathways. Need: earlier career-connected relevance. Opportunity: RISE Career & Math Mini Lessons.")
    if "Curriculum / HQIM" in tags or "Teacher Capacity" in tags: out.append("Priority: implementation quality. Need: teacher-ready routines after training. Opportunity: PCG implementation support and Emerald practice tools.")
    return (out or ["Priority: validate fit. Need: clarify district pain points and implementation constraints. Opportunity: discovery conversation."])[:3]

def relationship_insights(card):
    tags=card["tags"]
    out=["Build from existing relationship credibility before introducing new support." if "Existing Relationship" in tags else "Start with district priorities and implementation realities before discussing resources."]
    if "Curriculum / HQIM" in tags: out.append("Position PCG/Emerald as implementation support that strengthens adopted instructional materials.")
    elif "Teacher Capacity" in tags: out.append("Emphasize reducing teacher burden and strengthening routines already expected by the district.")
    else: out.append("Use consultative language around practical implementation and measurable proof points.")
    out.append("Best next step is a narrow proof point tied to a campus, grade band, or student group need.")
    return out

def compact(card):
    tags, signals = card["tags"], card["signals"]
    top = [t for t in tags if t not in ["Existing Relationship","Funding / Grants"]]
    why = f"High-interest fit around {', '.join(top[:3]).lower()}." if top else "Strategic fit should be validated through discovery."
    if card["priority"] in ["Very High","High"]:
        why = f"{card['priority']} priority account with signals around {', '.join(top[:3]).lower()}." if top else f"{card['priority']} priority account."
    barrier = "Likely barrier: campus variation, teacher capacity, or implementation consistency."
    if "Curriculum / HQIM" in tags: barrier = "Likely barrier: turning curriculum or HQIM expectations into consistent classroom routines."
    elif "MTSS" in tags: barrier = "Likely barrier: making intervention routines consistent across campuses."
    elif "SPED/ELL" in tags: barrier = "Likely barrier: maintaining grade-level access while differentiating support."
    return {"why":why,"entry":priority_need_opp(tags)[0],"barrier":barrier,"top_signals":[trunc(s,210) for s in signals[:3]],"positioning":priority_need_opp(tags),"relationship":relationship_insights(card)}

# -----------------------------
# Workbook load
# -----------------------------
def workbook_source(uploaded):
    if uploaded is not None: return uploaded, "Uploaded workbook override"
    if os.path.exists(BUILT_IN_WORKBOOK): return BUILT_IN_WORKBOOK, "Built-in workbook"
    return None, "No workbook found"

def load_cards(src):
    xls = pd.ExcelFile(src, engine="openpyxl")
    dfs = {name: read_sheet(xls, sheet) for name, sheet in SHEETS.items()}
    if dfs["strategic"].empty or dfs["scorecard"].empty:
        return [], xls, dfs, pd.DataFrame(), {}
    cols = score_cols(dfs["scorecard"])
    score_lookup = {}
    for _, r in dfs["scorecard"].iterrows():
        d = clean(r.get(cols["district"], "")) if cols["district"] else ""
        if d: score_lookup[dkey(d)] = r.to_dict()
    basic_lookup = lookup_by_district(dfs["basic"])
    contact_lookup = build_contacts_from_contacts(dfs["contacts"])
    leadership_lookup = build_contacts_from_leadership(dfs["leadership"])
    csi_counts, tsi_counts = counts_by_district(dfs["csi"]), counts_by_district(dfs["tsi"])
    cards, audit = [], []
    for _, strat in dfs["strategic"].iterrows():
        d = clean(val(strat, DISTRICT_COL))
        if not d: continue
        sr = score_lookup.get(dkey(d), {})
        matched = bool(sr)
        tier = clean(sr.get(cols.get("tier"), "")) if matched and cols.get("tier") else ""
        overall = sr.get(cols.get("overall"), "") if matched and cols.get("overall") else ""
        ok = substantive(strat) and matched and bool(tier) and bool(clean(overall))
        audit.append({"District": d, "Substantive Strategic Indicators": substantive(strat), "Matched Master Scorecard": matched, "Tier": tier, "Overall Score": clean(overall), "Eligible": ok})
        if not ok: continue
        tags = infer_tags(strat, sr)
        basic = basic_lookup.get(dkey(d), {})
        signals=[]
        for label, col in [("Strategic themes", "Strategic Plan Themes"),("Math signal","Math Priority Strength")]:
            if clean(val(strat,col)): signals.append(f"{label}: {clean(val(strat,col))}")
        csi = clean(val(basic,"CSI Schools")) or str(csi_counts.get(dkey(d), ""))
        tsi = clean(val(basic,"TSI Schools")) or str(tsi_counts.get(dkey(d), ""))
        if csi or tsi: signals.append(f"Accountability pressure: {csi or '0'} CSI schools and {tsi or '0'} TSI schools.")
        ctx=[]
        for lab, col, kind in [("schools","Number of Schools","text"),("grade span","Grade Span Served","text"),("setting","Urban/Suburban/Rural","text"),("economically disadvantaged","% Economically Disadvantaged","pct"),("English learner","% English Learner","pct"),("special education","% Special Education","pct"),("major student group","Major Student Groups","text"),("student growth trend","Student Growth Trend","pct1")]:
            raw = val(basic,col); v=clean(raw)
            if v:
                if kind=="pct": v=fmt_pct(raw,0)
                elif kind=="pct1": v=fmt_pct(raw,1)
                ctx.append(f"{lab}: {v}")
        if ctx: signals.append("District context: " + "; ".join(ctx) + ".")
        for lab, col in [("Intervention signal","Intervention Focus Details"),("MTSS signal","MTSS Details"),("SPED/ELL signal","SPED/ELL Details"),("Teacher capacity signal","Teacher Capacity Details"),("CCMR / career readiness signal","Career Readiness Details"),("Curriculum / implementation signal","Curriculum Details"),("Funding / grants signal","Grants Details")]:
            if clean(val(strat,col)): signals.append(f"{lab}: {clean(val(strat,col))}")
        rel, con = clean(val(sr,"Existing Relationships")), clean(val(sr,"Existing Contracts in Region (Yes/No)"))
        if rel or con: signals.append(f"Relationship context: existing contracts in region = {con or 'Unknown'}; existing relationships = {rel or 'Unknown'}.")
        card = {"name":d,"tier":tier,"score":fmt_num(overall,2),"strategic_score":fmt_num(sr.get(cols.get("strategic"), ""),2) if cols.get("strategic") else "","enrollment":fmt_enroll(sr.get(cols.get("enrollment"),"")) if cols.get("enrollment") else "","priority":priority(tier, overall),"tags":tags,"contacts":get_contacts(d, contact_lookup, leadership_lookup),"signals":signals,"alignment":alignment(tags),"listen":listen_for(tags),"engagement_guidance":guidance(tags)}
        card["lead"] = build_lead_with(card); card["questions"] = questions(tags); card["compact"] = compact(card)
        cards.append(card)
    return cards, xls, dfs, pd.DataFrame(audit), cols

# -----------------------------
# Exports
# -----------------------------
def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), fill); tc_pr.append(shd)

def add_bullets(cell, items, limit=None):
    cell.text = ""
    for item in (items[:limit] if limit else items):
        if clean(item): cell.add_paragraph(f"• {item}")

def build_docx(cards):
    doc=Document(); title=doc.add_heading("Strategic District Field Guide",0)
    for r in title.runs: r.font.color.rgb=RGBColor(23,54,93)
    doc.add_paragraph("Mobile conversation cards for conference prep: strategic signals, contacts, PCG/Emerald alignment, and discovery questions.")
    for c in cards:
        comp=c["compact"]; doc.add_heading(c["name"],1); doc.add_paragraph(f"{c['tier']} | Overall Score {c['score']} | Strategic Score {c['strategic_score']} | Enrollment {c['enrollment']} | Priority {c['priority']}")
        doc.add_heading("Quick Read",2)
        for item in [comp["why"], comp["entry"], comp["barrier"]]: doc.add_paragraph(item, style="List Bullet")
        for title, field in [("Top Strategic Signals","signals"),("Best-Fit PCG / Emerald Alignment","alignment"),("Key Contacts","contacts"),("Discovery Questions","questions"),("Listen For","listen"),("PCG Engagement Guidance","engagement_guidance")]:
            doc.add_heading(title,2)
            for item in c.get(field,[]): doc.add_paragraph(item, style="List Bullet")
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio

def build_matrix_docx(cards):
    doc=Document(); sec=doc.sections[0]; sec.left_margin=Inches(.35); sec.right_margin=Inches(.35); sec.top_margin=Inches(.4); sec.bottom_margin=Inches(.4)
    title=doc.add_heading("Texas District Strategic Positioning Matrix",0)
    for r in title.runs: r.font.color.rgb=RGBColor(31,78,121)
    doc.add_paragraph("Scope: Currently filtered districts from the Strategic District Field Guide.")
    doc.add_paragraph("Designed for deeper review after quick conference scanning in the app.")
    headers=["District / Score / Key Contacts","District Overview","Strategic Signal Analysis","Positioning Matrix","Discovery Questions","Relationship & Engagement Insights"]
    table=doc.add_table(rows=1, cols=6); table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.style="Table Grid"
    for i,h in enumerate(headers):
        cell=table.rows[0].cells[i]; cell.text=h; shade(cell,"1F4E79")
        for p in cell.paragraphs:
            for r in p.runs: r.font.color.rgb=RGBColor(255,255,255); r.bold=True
    for c in cards:
        comp=c["compact"]; cells=table.add_row().cells
        for cell in cells: cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
        cells[0].text=f"{c['name']}\n{c['tier']} | Overall {c['score']}\nEnrollment: {c['enrollment']}\nContacts: " + "; ".join(c.get("contacts",[])[:5])
        add_bullets(cells[1],[comp["why"],comp["entry"],comp["barrier"]]); add_bullets(cells[2],comp["top_signals"],3); add_bullets(cells[3],comp["positioning"],3); add_bullets(cells[4],c["questions"],6); add_bullets(cells[5],comp["relationship"],3)
    doc.add_heading("PCG / Emerald Alignment Lens",1)
    lens=doc.add_table(rows=1, cols=3); lens.style="Table Grid"; lh=["Strategic Need","Consultative Alignment","Representative Resource / Service Fit"]
    for i,h in enumerate(lh):
        cell=lens.rows[0].cells[i]; cell.text=h; shade(cell,"1F4E79")
        for p in cell.paragraphs:
            for r in p.runs: r.font.color.rgb=RGBColor(255,255,255); r.bold=True
    rows=[("Math proficiency / early numeracy / Algebra readiness","Implementation support, progress monitoring routines, supplemental practice, teacher-facing supports","Elevation Station Math Games; Elevation intervention curriculum; PCG instructional implementation support"),("MTSS / Tier 2 / Tier 3 intervention","System design, campus implementation, intervention fidelity, data-to-action routines","PCG MTSS consulting; Emerald intervention supports"),("SPED / ELL / subgroup performance","Inclusive practice, service delivery, compliance-to-instruction alignment, differentiated supports","PCG SPED and multilingual learner services; scaffolded Emerald resources"),("CCMR / career-connected learning","Middle-grade pathway exposure, career relevance, readiness pipeline alignment","RISE Career & Math Mini Lessons; PCG CCMR strategy support"),("Teacher capacity / HQIM or curriculum adoption","Professional learning, coaching, PLC routines, adoption fidelity, change management","PCG implementation and professional learning services; Emerald teacher-ready practice tools")]
    for a,b,c in rows:
        cells=lens.add_row().cells; cells[0].text=a; cells[1].text=b; cells[2].text=c
    doc.add_paragraph("Note: This matrix is designed for strategy review and conference/meeting preparation. Validate against current district conversations before final pursuit decisions.")
    bio=BytesIO(); doc.save(bio); bio.seek(0); return bio

# -----------------------------
# UI
# -----------------------------
def show_diag(xls, dfs, audit, cols, source_label):
    st.markdown('<div class="section-title">Workbook Diagnostics</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="helper-text">Current data source: <strong>{safe_html(source_label)}</strong></div>', unsafe_allow_html=True)
    st.write(xls.sheet_names); st.write(cols)
    for label, df in dfs.items():
        with st.expander(label, expanded=False):
            st.write(f"Rows: {len(df)}"); st.write(list(df.columns))
            if DISTRICT_COL in df.columns: st.write(df[DISTRICT_COL].dropna().astype(str).head(15).tolist())
    st.dataframe(audit, use_container_width=True)

def render_howto(source_label):
    st.markdown('<div class="section-title">How-To Guide</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="howto-box"><strong>Current data source:</strong> {safe_html(source_label)}<br>The app opens with the built-in workbook from GitHub. Uploading a workbook temporarily overrides the built-in data for that session.</div>', unsafe_allow_html=True)
    st.markdown('<div class="howto-box"><strong>60-second workflow</strong><ol><li>Search the district name.</li><li>Read Why It Matters, Best Entry Point, and Likely Barrier.</li><li>Use the Quick Prep question to start discovery.</li><li>Open Key Contacts for names, titles, and emails.</li><li>Use the Full Matrix download for deeper planning.</li></ol></div>', unsafe_allow_html=True)
    st.markdown("### Useful searches"); st.code("Dallas\nFort Worth\nSPED\nMTSS\nEureka\nBluebonnet\ncareer\nteacher burden", language="text")

def render_matrix(cards):
    st.markdown('<div class="section-title">Matrix View</div>', unsafe_allow_html=True)
    if not cards: st.info("No cards to show."); return
    df=pd.DataFrame([{"District":c["name"],"Tier":c["tier"],"Score":c["score"],"Why It Matters":c["compact"]["why"],"Best Entry Point":c["compact"]["entry"],"Best Question":(c["questions"] or [""])[0],"Contacts":"; ".join(c.get("contacts",[])[:3])} for c in cards])
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Download short matrix as CSV", df.to_csv(index=False).encode("utf-8"), "district_opportunity_matrix.csv", "text/csv")

def render_card(c, view_mode):
    pc=c["priority"].lower().replace(" ","-"); badges="".join(badge(t) for t in c["tags"]); comp=c["compact"]
    st.markdown(f'<div class="card"><h3>{safe_html(c["name"])} <span class="priority priority-{pc}">{safe_html(c["priority"])}</span></h3><div class="meta">{safe_html(c["tier"])} | Overall Score {safe_html(c["score"])} | Strategic Score {safe_html(c["strategic_score"])} | Enrollment {safe_html(c["enrollment"])}</div><div>{badges}</div></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="summary-grid"><div class="summary-box"><div class="summary-label">Why it matters</div><div class="summary-value">{safe_html(comp["why"])}</div></div><div class="summary-box"><div class="summary-label">Best entry point</div><div class="summary-value">{safe_html(comp["entry"])}</div></div><div class="summary-box"><div class="summary-label">Likely barrier</div><div class="summary-value">{safe_html(comp["barrier"])}</div></div></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="quickprep"><div class="quickprep-title">Quick Prep</div><strong>Best opening question:</strong> {safe_html(c["questions"][0] if c["questions"] else "")}<br><strong>Listen for:</strong><br>{"".join(chip(x) for x in c["listen"][:8])}<br><strong>PCG engagement note:</strong> {safe_html(c["engagement_guidance"][0] if c["engagement_guidance"] else "")}</div>', unsafe_allow_html=True)
    st.markdown("**Top Signals**"); [st.markdown(f"- {x}") for x in comp["top_signals"][:3]]
    st.markdown("**Best Fit**"); [st.markdown(f"- {x}") for x in comp["positioning"][:3]]
    st.markdown("**Discovery Questions**"); [st.markdown(f"- {x}") for x in c["questions"][:4]]
    with st.expander("Key Contacts", expanded=False):
        if c.get("contacts"): [st.markdown(f"- {x}") for x in c["contacts"][:6]]
        else: st.markdown("_No contacts were available for this district._")
    if view_mode == "Full Detail":
        with st.expander("Full Strategic Detail", expanded=False):
            st.markdown("**Strategic Signals**"); [st.markdown(f"- {x}") for x in c["signals"]]
            st.markdown("**PCG / Emerald Alignment**"); [st.markdown(f"- {x}") for x in c["alignment"]]
            st.markdown("**Listen For**"); st.markdown("".join(chip(x) for x in c["listen"]), unsafe_allow_html=True)
            st.markdown("**PCG Engagement Guidance**"); [st.markdown(f"- {x}") for x in c["engagement_guidance"]]
    copy = f"{c['name']} Quick Prep\n\nWhy it matters: {comp['why']}\nBest entry point: {comp['entry']}\nLikely barrier: {comp['barrier']}\n\nBest opening question: {(c['questions'] or [''])[0]}\n\nListen for: {', '.join(c['listen'][:8])}\n\nPCG engagement note: {(c['engagement_guidance'] or [''])[0]}"
    with st.expander("Copy Quick Prep", expanded=False): st.text_area("Quick prep copy", copy, height=150, key=f"quick_{c['name']}")
    st.download_button("Download this district brief", build_docx([c]), f"{c['name'].replace(' ', '_')}_Brief.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"download_{c['name']}")

st.markdown('<div class="main-title">Strategic District Field Guide</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Quick district refresh cards for conferences, plus deeper matrix exports for follow-up planning.</div>', unsafe_allow_html=True)
with st.sidebar:
    st.header("1. Data"); uploaded = st.file_uploader("Optional: upload workbook override", type=["xlsx"]); st.caption("If no workbook is uploaded, the app uses the built-in workbook from GitHub.")
src, source_label = workbook_source(uploaded)
if src is None:
    tabs = st.tabs(["Field Guide", "How-To Guide", "Matrix View", "Workbook Diagnostics"])
    with tabs[0]: st.error(f"No workbook found. Add {BUILT_IN_WORKBOOK} to the GitHub repo or upload a workbook in the sidebar.")
    with tabs[1]: render_howto(source_label)
    with tabs[2]: st.info("Workbook required to populate the matrix.")
    with tabs[3]: st.info("Workbook required to view diagnostics.")
    st.stop()

cards, xls, dfs, audit, cols = load_cards(src)
if not cards:
    tabs = st.tabs(["Field Guide", "How-To Guide", "Matrix View", "Workbook Diagnostics"])
    with tabs[0]: st.warning("No eligible district cards were generated. Review Workbook Diagnostics.")
    with tabs[1]: render_howto(source_label)
    with tabs[2]: st.info("No eligible cards available.")
    with tabs[3]: show_diag(xls, dfs, audit, cols, source_label)
    st.stop()

all_districts=[c["name"] for c in cards]; all_tiers=sorted({c["tier"] for c in cards if c["tier"]}); all_tags=sorted({t for c in cards for t in c["tags"]})
with st.sidebar:
    st.success(f"Using: {source_label}")
    st.header("2. Search"); query=st.text_input("Search", placeholder="District, contact, signal, offering...")
    st.header("3. Filters"); selected_tiers=st.multiselect("Tier", all_tiers, default=[]); selected_tags=st.multiselect("Strategic / solution tags", all_tags, default=[])
    st.header("4. Shortlist"); shortlist=st.multiselect("Shortlist districts", all_districts, default=[]); show_short=st.checkbox("Show shortlisted only", False)
    st.header("5. Display"); view_mode=st.radio("View mode", ["Quick Brief","Full Detail"], index=0); st.caption("Tip: For conferences, use Quick Brief. For prep or follow-up, download the full matrix.")

def blob(c): return " ".join(map(str, [c.get("name",""), c.get("tier",""), c.get("score",""), c.get("lead","")] + c.get("tags",[]) + c.get("signals",[]) + c.get("contacts",[]) + c.get("questions",[]) + c.get("listen",[]) + c.get("engagement_guidance",[]))).lower()
filtered=[]
for c in cards:
    if selected_tiers and c["tier"] not in selected_tiers: continue
    if selected_tags and not any(t in c["tags"] for t in selected_tags): continue
    if show_short and c["name"] not in shortlist: continue
    if query and query.lower().strip() not in blob(c): continue
    filtered.append(c)

tab_field, tab_howto, tab_matrix, tab_diag = st.tabs(["Field Guide", "How-To Guide", "Matrix View", "Workbook Diagnostics"])
with tab_field:
    m1,m2,m3,m4=st.columns(4); m1.metric("📍 Cards Shown", len(filtered)); m2.metric("✅ Eligible", len(cards)); m3.metric("🔥 High Priority", sum(1 for c in filtered if c["priority"] in ["High","Very High"])); m4.metric("⭐ Tier 1", sum(1 for c in filtered if c["tier"]=="Tier 1"))
    if query: st.markdown(f'<div class="metric-note">Showing results for: <strong>{safe_html(query)}</strong></div>', unsafe_allow_html=True)
    d1,d2,d3=st.columns([1,1,1])
    with d1: st.download_button("Download brief Word guide", build_docx(filtered), "Strategic_District_Field_Guide.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", disabled=not filtered)
    with d2: st.download_button("Download full matrix", build_matrix_docx(filtered), "Texas_District_Strategic_Positioning_Matrix.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", disabled=not filtered)
    with d3:
        sel=st.selectbox("One-district brief", [""]+[c["name"] for c in filtered])
        if sel:
            sc=next(c for c in filtered if c["name"]==sel)
            st.download_button("Download selected brief", build_docx([sc]), f"{sel.replace(' ', '_')}_Brief.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="download_selected_brief")
    st.divider()
    if not filtered: st.info("No district cards match the current search/filter.")
    for c in filtered:
        render_card(c, view_mode); st.divider()
with tab_howto: render_howto(source_label)
with tab_matrix:
    render_matrix(filtered)
    st.download_button("Download full matrix as Word document", build_matrix_docx(filtered), "Texas_District_Strategic_Positioning_Matrix.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", disabled=not filtered, key="matrix_tab_download")
with tab_diag: show_diag(xls, dfs, audit, cols, source_label)
