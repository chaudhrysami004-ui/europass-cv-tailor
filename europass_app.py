import os
import io
import re
import time
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import streamlit as st
from google import genai

st.set_page_config(
    page_title="Shafay Munir — Autonomous ATS Tailor",
    page_icon="⚡",
    layout="wide"
)

# Ultra High-Visibility Light-Text Theme with Ambient Motion
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

    * {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    /* Ambient Animated Mesh Background */
    .stApp {
        background: linear-gradient(-45deg, #070B14, #0F172A, #172554, #0A0F1D) !important;
        background-size: 300% 300% !important;
        animation: cyberFlow 18s ease infinite !important;
    }

    @keyframes cyberFlow {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Subtle Grid Pattern Overlay */
    .stApp::before {
        content: "";
        position: fixed;
        top: 0; left: 0; width: 100vw; height: 100vh;
        background-image: linear-gradient(rgba(147, 197, 253, 0.05) 1px, transparent 1px),
                          linear-gradient(90deg, rgba(147, 197, 253, 0.05) 1px, transparent 1px);
        background-size: 32px 32px;
        pointer-events: none;
        z-index: 0;
    }

    /* Hero Header */
    .hero-header {
        position: relative;
        padding: 1.8rem 2.2rem;
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 16px;
        backdrop-filter: blur(16px);
        margin-bottom: 2rem;
        box-shadow: 0 15px 35px -10px rgba(0, 0, 0, 0.6);
        overflow: hidden;
    }

    .hero-header::before {
        content: '';
        position: absolute;
        top: 0; left: -100%; width: 100%; height: 3px;
        background: linear-gradient(90deg, transparent, #818CF8, #38BDF8, transparent);
        animation: sweep 3.5s linear infinite;
    }

    @keyframes sweep {
        0% { left: -100%; }
        100% { left: 100%; }
    }

    .hero-title {
        font-size: 2.2rem;
        font-weight: 800;
        color: #FFFFFF !important;
        margin: 0;
        letter-spacing: -0.02em;
        text-shadow: 0 0 25px rgba(99, 102, 241, 0.5);
    }

    .hero-desc {
        color: #E2E8F0 !important;
        font-size: 0.98rem;
        margin-top: 0.4rem;
        font-weight: 500;
    }

    .status-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.35rem 0.9rem;
        background: rgba(99, 102, 241, 0.25);
        border: 1px solid #818CF8;
        border-radius: 9999px;
        color: #E0E7FF !important;
        font-size: 0.78rem;
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
        margin-top: 0.8rem;
        box-shadow: 0 0 12px rgba(99, 102, 241, 0.3);
    }

    /* ALL LABELS FORCED TO BRIGHT LIGHT COLOR */
    label, .stTextInput label, .stTextArea label, .stSelectbox label, p, span, h1, h2, h3 {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        text-shadow: 0 1px 3px rgba(0,0,0,0.8);
    }

    /* Subheadings High Contrast */
    h3 {
        color: #F8FAFC !important;
        font-weight: 800 !important;
        letter-spacing: -0.01em;
    }

    /* HIGH-CONTRAST INPUT BOXES */
    .stTextInput input, .stTextArea textarea {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        border: 1.5px solid #475569 !important;
        border-radius: 10px !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.95rem !important;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.4) !important;
    }

    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #38BDF8 !important;
        box-shadow: 0 0 16px rgba(56, 189, 248, 0.5) !important;
        background-color: #0B1120 !important;
    }

    /* BRIGHT PLACEHOLDERS */
    ::placeholder {
        color: #94A3B8 !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #94A3B8 !important;
        font-size: 0.92rem !important;
    }

    /* Selectbox dropdown fix */
    div[data-baseweb="select"] {
        background-color: #0F172A !important;
        border: 1.5px solid #475569 !important;
        border-radius: 10px !important;
    }

    div[data-baseweb="select"] * {
        background-color: transparent !important;
        color: #FFFFFF !important;
    }

    /* Context Panel Card */
    .context-panel {
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid rgba(147, 197, 253, 0.25);
        border-radius: 12px;
        padding: 1.2rem;
        margin-top: 1.4rem;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
    }

    .context-panel-title {
        color: #38BDF8 !important;
        font-weight: 800;
        font-size: 0.9rem;
        letter-spacing: 0.05em;
        margin-bottom: 0.5rem;
    }

    .context-panel-body {
        font-size: 0.9rem;
        color: #E2E8F0 !important;
        line-height: 1.7;
    }

    /* Glowing Compile Button */
    .stButton > button {
        background: linear-gradient(135deg, #4F46E5 0%, #6366F1 50%, #38BDF8 100%) !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 1.05rem !important;
        padding: 0.9rem 1.8rem !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 6px 24px rgba(99, 102, 241, 0.6) !important;
        transition: all 0.25s ease !important;
        width: 100% !important;
        margin-top: 1.2rem !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px) scale(1.01) !important;
        box-shadow: 0 10px 32px rgba(56, 189, 248, 0.8) !important;
    }

    /* Download File Button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%) !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 800 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 0.9rem 1.6rem !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.5) !important;
    }
</style>
""", unsafe_allow_html=True)

# Master Profiles
PROFILES = {
    "AI SEO & GEO Specialist": """
CANDIDATE: Shafay Munir
ROLE: AI SEO / GEO Specialist
LOCATION: Vilnius, Lithuania (TRP Valid until 2029 - Full legal right to work)
CONTACT: +370 692 28 728 | shafaymunir890@gmail.com | linkedin.com/in/shafay-munir/
EDUCATION:
- MSc in International Marketing & Management, ISM University of Management and Economics, Vilnius, Lithuania (2026 - Present)
- BSc in Information Technology, Bahria University, Islamabad, Pakistan (2019 - 2023)
EXPERIENCE:
1. AI SEO Specialist | AIO (Silicon Valley, San Jose, CA · Hybrid/Remote) [Aug 2024 - Jul 2026]
- Joined in stealth mode with 0 clients; scaled platform to 50+ onboarded restaurant clients and through a $17M Series A raise as primary SEO owner.
- Led end-to-end website architecture and structural planning for SEO & CRO, taking aioapp.com from zero to 1.02M search impressions and 21.7K clicks within 12 months.
- Designed and executed 3-layer SEO research framework (search intent, page placement, content angle) for all product and marketing page builds.
- Implemented semantic SEO and demo-led frameworks to strengthen topical authority and search intent alignment.
- Conducted biweekly AI visibility audits across ChatGPT, Perplexity, Gemini, and Google AI Overviews; tracked citation frequency, entity signals, and share of voice.
- Spearheaded SEO-driven content automation via NLP-based models, reducing production time from hours to minutes at scale.
- Conducted technical SEO audits covering crawlability, indexation, structured data, schema markup, and Core Web Vitals.
2. SEO Team Lead | Cretesol Technologies (Dubai, UAE HQ) [Mar 2024 - Aug 2024]
- Led end-to-end SEO strategy and technical audits across client sites; managed and mentored a team of SEO specialists.
- Directed on-page, advanced keyword research, and high-quality link-building campaigns.
3. SEO Executive | United Sol [Sep 2023 - Mar 2024]
- Achieved top-3 US rankings for 9 Shopify apps; increased domain authority to DA 59 through strategic outreach.
CORE TOOLS: Google Search Console, Google Analytics 4, Ahrefs, Semrush, Screaming Frog, Sitebulb, SimilarWeb, Surfer SEO, GTM, n8n, ChatGPT, Perplexity, Gemini, Claude, WordPress, Shopify.
""",

    "Associate Product Manager (APM)": """
CANDIDATE: Shafay Munir
ROLE: Associate Product Manager
LOCATION: Vilnius, Lithuania (TRP Valid until 2029 - Full legal right to work)
CONTACT: +370 692 28 728 | shafaymunir890@gmail.com | linkedin.com/in/shafay-munir/
EDUCATION:
- MSc in International Marketing & Management, ISM University of Management and Economics, Vilnius, Lithuania (2026 - Present)
- BSc in Information Technology, Bahria University, Islamabad, Pakistan (2019 - 2023)
EXPERIENCE:
1. Associate Product Manager | AIO (Silicon Valley, San Jose, CA) [Aug 2024 - Jul 2026]
- Joined as early team member in stealth mode; contributed to scaling from 0 to 50+ clients and $17M Series A raise, owning the AIO Marketing module.
- End-to-end product owner for AIO Website Builder (primary onboarding entry point across 50+ live restaurant sites); defined requirements for template system, ADA compliance, reservations, and online ordering.
- Owned product requirements for Smart CRM (automated guest profile builds from order history, behavioral segmentation, automated campaigns).
- Led product definition for Automated Reputation Management (ARM), cutting manual review management time by ~70% via AI-generated sentiment responses.
- Designed & spec'd 8-module Content Automation product system (voice selection, outline editor, tone configuration, length control, blog generator).
- Defined Marketing Overview Dashboard (Growth, SEO, ADA tabs) tracking CTR, traffic trends, and site health across all client sites.
SKILLS: Product Strategy & Roadmapping, Feature Prioritization, User Story Definition, Go-to-Market Execution, PRDs, KPI Framework Design, SaaS Platform Ownership, Cross-functional Leadership.
""",

    "General Services & Operations (Warehouse / Hospitality / Retail)": """
CANDIDATE: Shafay Munir
ROLE: Team Member / Operations & Service Assistant
LOCATION: Vilnius, Lithuania (TRP Valid until 2029 - Available full-time / immediate start)
CONTACT: +370 692 28 728 | shafaymunir890@gmail.com
SUMMARY:
Hardworking, physically fit, and dependable professional based in Vilnius. Punctual, fast learner, and experienced in high-demand team environments. Available for immediate full-time work across shift rotations (morning, evening, night, weekends).
WORK EXPERIENCE:
1. Operations Team Member | AIO [Aug 2024 - Jul 2026]
- Worked in a fast-paced environment managing high-priority deliverables under tight deadlines.
- Maintained strict operational quality control and cross-functional task coordination.
2. Team Coordinator | Cretesol Technologies [Mar 2024 - Aug 2024]
- Coordinated team schedules, daily task prioritization, and operational workflow delivery.
3. Operations Assistant | United Sol [Sep 2023 - Mar 2024]
- Handled day-to-day structured execution, inventory of digital tasks, and routine reporting.
KEY STRENGTHS:
- Punctual, reliable, high physical stamina, comfortable with manual and fast-paced shift work.
- Fast learner, adaptable, team-oriented, strict safety and procedural compliance.
- Languages: English (Fluent/Professional), Urdu (Native), Lithuanian (Basic/A0 - learning).
"""
}

def generate_tailored_cv(profile_name, job_desc, api_key, custom_model_name):
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
You are an expert European ATS Resume Writer specializing in the Baltic/Lithuanian and EU job market.
Your task is to tailor an authentic, ATS-compliant Europass layout CV for candidate Shafay Munir based strictly on his master background and the target Job Description.

MASTER DATA:
{PROFILES[profile_name]}

TARGET JOB DESCRIPTION:
{job_desc}

RULES & GUIDELINES:
1. Tone & Culture: Factual, professional, metric-driven (for tech roles) or reliability-focused (for operations/services).
2. Truthfulness: Strictly stick to the master background numbers (e.g. 50+ clients, 1.02M impressions, $17M Series A). Do not invent fake companies or credentials.
3. CRITICAL EUROPASS FORMAT REQUIREMENT:
   Every Experience and Education entry MUST be formatted with '## ' followed by: [Role/Degree] | [Company/University] [[Dates]]
   Example:
   ## AI SEO Specialist | AIO (Silicon Valley, CA) [Aug 2024 - Jul 2026]
   - Bullet point achievement 1
   - Bullet point achievement 2

OUTPUT STRUCTURE (Strictly follow without preamble, conversational remarks, or code fences):

# SHAFAY MUNIR
[Target Job Title] | Vilnius, Lithuania | +370 692 28 728 | shafaymunir890@gmail.com | linkedin.com/in/shafay-munir/

# PROFESSIONAL SUMMARY
(3-4 impactful lines targeted to the JD embedding keywords)

# WORK EXPERIENCE
## [Job Title] | [Company, Location] [Start Date - End Date]
- Impact bullet point 1
- Impact bullet point 2
- Impact bullet point 3

# EDUCATION AND TRAINING
## [Degree Title] | [University, Location] [Start Date - End Date]
- Relevant focus / achievements

# DIGITAL & CORE SKILLS
- Skill category 1: details
- Skill category 2: details

# TOOLS & TECHNOLOGIES
- Tools list

# LANGUAGE SKILLS
- Mother tongue: Urdu
- Other languages: English (Proficient / C1-C2), Lithuanian (Basic / A1)
"""
    candidate_models = [
        custom_model_name.strip(),
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-flash-latest"
    ]
    candidate_models = [m for m in dict.fromkeys(candidate_models) if m]

    last_err = None
    for m in candidate_models:
        for _ in range(2):
            try:
                response = client.models.generate_content(
                    model=m,
                    contents=prompt
                )
                if response and response.text:
                    return response.text
            except Exception as err:
                last_err = err
                time.sleep(1.2)
                continue

    raise last_err

def set_cell_border_none(cell):
    """Europass table grid borders remove karne ke liye"""
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'none')
        tcBorders.append(b)
    tcPr.append(tcBorders)

def create_europass_docx(markdown_content):
    doc = Document()
    
    # Official Page Setup (A4 Margins standard for Europass)
    for s in doc.sections:
        s.top_margin = Inches(0.55)
        s.bottom_margin = Inches(0.55)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.7)

    COLOR_PRIMARY_BLUE = RGBColor(14, 65, 148)    # Europass Official Navy Blue (#0E4194)
    COLOR_BODY = RGBColor(45, 55, 72)             # Deep Charcoal
    COLOR_MUTED = RGBColor(100, 116, 139)         # Timeline Grey

    lines = markdown_content.strip().split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # MAIN SECTION HEADINGS (e.g. # WORK EXPERIENCE, # EDUCATION AND TRAINING)
        if line.startswith("# "):
            title = line.replace("# ", "").strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(title.upper())
            run.font.name = "Arial"
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = COLOR_PRIMARY_BLUE
            
            # Europass bottom accent divider line under headings
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), '0E4194')
            pBdr.append(bottom)
            p._p.get_or_add_pPr().append(pBdr)
            i += 1

        # EXPERIENCE / EDUCATION 2-COLUMN TIMELINE BLOCKS (Official Europass Layout)
        elif line.startswith("## "):
            header_text = line.replace("## ", "").strip()
            
            date_part = ""
            main_part = header_text
            
            # Extract bracketed dates [Aug 2024 - Jul 2026]
            date_match = re.search(r'\[(.*?)\]', header_text)
            if date_match:
                date_part = date_match.group(1)
                main_part = header_text.replace(f"[{date_part}]", "").strip(" | -")
            
            # Europass 2-Column Table Grid
            table = doc.add_table(rows=1, cols=2)
            table.autofit = False
            
            cell_left = table.cell(0, 0)
            cell_right = table.cell(0, 1)
            cell_left.width = Inches(1.8)   # Left Timeline
            cell_right.width = Inches(5.2)  # Right Details
            
            set_cell_border_none(cell_left)
            set_cell_border_none(cell_right)

            # Left Cell: Timeline / Period
            p_date = cell_left.paragraphs[0]
            p_date.paragraph_format.space_before = Pt(2)
            p_date.paragraph_format.space_after = Pt(2)
            r_date = p_date.add_run(date_part)
            r_date.font.name = "Arial"
            r_date.font.size = Pt(9)
            r_date.font.bold = True
            r_date.font.color.rgb = COLOR_MUTED

            # Right Cell: Role, Company & Location
            p_role = cell_right.paragraphs[0]
            p_role.paragraph_format.space_before = Pt(2)
            p_role.paragraph_format.space_after = Pt(3)
            r_role = p_role.add_run(main_part)
            r_role.font.name = "Arial"
            r_role.font.size = Pt(10)
            r_role.font.bold = True
            r_role.font.color.rgb = COLOR_PRIMARY_BLUE

            # Subsequent bullet points ko right cell ke andar render karein
            i += 1
            while i < len(lines) and (lines[i].strip().startswith("- ") or lines[i].strip().startswith("* ")):
                b_text = lines[i].strip()[2:].strip()
                p_bullet = cell_right.add_paragraph(style='List Bullet')
                p_bullet.paragraph_format.space_before = Pt(1)
                p_bullet.paragraph_format.space_after = Pt(1.5)
                
                parts = re.split(r'(\*\*.*?\*\*)', b_text)
                for part in parts:
                    if part.startswith("**") and part.endswith("**"):
                        r = p_bullet.add_run(part[2:-2])
                        r.bold = True
                    else:
                        r = p_bullet.add_run(part)
                    r.font.name = "Arial"
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = COLOR_BODY
                i += 1

        # REGULAR BULLET POINTS (Skills, Languages, Tools)
        elif line.startswith("- ") or line.startswith("* "):
            bullet_text = line[2:].strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1.5)
            
            parts = re.split(r'(\*\*.*?\*\*)', bullet_text)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = "Arial"
                r.font.size = Pt(9.5)
                r.font.color.rgb = COLOR_BODY
            i += 1

        # REGULAR PARAGRAPHS (Header Info, Summary)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = "Arial"
                r.font.size = Pt(9.5)
                r.font.color.rgb = COLOR_BODY
            i += 1

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Hero Banner
st.markdown("""
<div class="hero-header">
    <div class="hero-title">⚡ Shafay Munir — Autonomous ATS Tailor</div>
    <div class="hero-desc">Real-time reverse-chronological Europass CV generation calibrated for the Baltic & EU market.</div>
    <div class="status-chip">● ENGINE ACTIVE: GEMINI 2.0 FLASH / AUTO-HEALING</div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1.1, 1.9], gap="large")

with col1:
    st.markdown("### ⚙️ Pipeline Configuration")
    api_key = st.text_input("Gemini API Key", type="password", help="Personal AI Studio API Key")
    model_choice = st.text_input("Model Engine ID", value="gemini-2.0-flash")
    target_track = st.selectbox("Select Target Track", list(PROFILES.keys()))
    
    st.markdown("""
    <div class="context-panel">
        <div class="context-panel-title">PRESET CONTEXT LOADED</div>
        <div class="context-panel-body">
            • 📍 <b>Vilnius, Lithuania</b> (TRP Valid until 2029)<br>
            • 🎓 <b>ISM University MSc</b> & Bahria BSc<br>
            • 🚀 <b>AIO</b> ($17M Series A, 0 to 50 scale, 1M+ views)
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("### 🎯 Target Job Description")
    job_desc = st.text_area(
        "Paste Job Description text here:",
        height=260,
        placeholder="Paste JD requirements, skills, and duties here..."
    )
    
    generate_trigger = st.button("🚀 Compile Tailored Europass CV")

if generate_trigger:
    if not api_key.strip():
        st.error("Pehle Gemini API key enter karein.")
    elif not job_desc.strip():
        st.error("Job description paste karna zaroori hai.")
    else:
        with st.status("⚡ Initializing Autonomous ATS Pipeline...", expanded=True) as status:
            st.write("🔍 Parsing Job Description intent and high-frequency ATS tokens...")
            time.sleep(0.3)
            st.write(f"🧠 Synthesizing Master Background for track: **{target_track}**...")
            try:
                cv_markdown = generate_tailored_cv(target_track, job_desc, api_key, model_choice)
                st.write("📄 Structuring Europass Word XML document...")
                docx_file = create_europass_docx(cv_markdown)
                status.update(label="✅ Compilation Complete!", state="complete", expanded=False)
                
                clean_name = target_track.split()[0]
                
                st.success("✔ Document compiled successfully!")
                st.download_button(
                    label=f"📥 Download Tailored CV ({clean_name}.docx)",
                    data=docx_file,
                    file_name=f"Shafay_Munir_CV_{clean_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                with st.expander("👁️ View Live Markdown Preview"):
                    st.markdown(cv_markdown)
                    
            except Exception as e:
                status.update(label="❌ Pipeline Failed", state="error", expanded=True)
                st.error(f"Execution Error: {e}")
