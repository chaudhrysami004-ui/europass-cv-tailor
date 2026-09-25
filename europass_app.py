import os
import io
import re
import time
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

    /* High Visibility Labels and Texts */
    label, .stTextInput label, .stTextArea label, .stSelectbox label, .stRadio label, p, span, h1, h2, h3 {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 0.98rem !important;
        text-shadow: 0 1px 3px rgba(0,0,0,0.8);
    }

    h3 {
        color: #F8FAFC !important;
        font-weight: 800 !important;
        letter-spacing: -0.01em;
    }

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

    ::placeholder {
        color: #94A3B8 !important;
        opacity: 1 !important;
        -webkit-text-fill-color: #94A3B8 !important;
        font-size: 0.92rem !important;
    }

    div[data-baseweb="select"] {
        background-color: #0F172A !important;
        border: 1.5px solid #475569 !important;
        border-radius: 10px !important;
    }

    div[data-baseweb="select"] * {
        background-color: transparent !important;
        color: #FFFFFF !important;
    }

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

# 10 Major EU & Target Languages
LANGUAGES_AVAILABLE = [
    "English (Default)",
    "Lithuanian (Lietuvių k.)",
    "German (Deutsch)",
    "Polish (Polski)",
    "French (Français)",
    "Dutch (Nederlands)",
    "Latvian (Latviešu)",
    "Estonian (Eesti)",
    "Spanish (Español)",
    "Italian (Italiano)"
]

def generate_tailored_cv(profile_name, job_desc, api_key, custom_model_name, target_lang, trans_scope):
    client = genai.Client(api_key=api_key)

    lang_instructions = f"""
LANGUAGE & LOCALIZATION REQUIREMENTS:
- Target Output Language: {target_lang}
- Translation Scope: {trans_scope}
"""
    if trans_scope == "Entire CV":
        lang_instructions += f"""
- Translate all section headers, summaries, experience bullet points, skills, and dates into {target_lang}.
- Keep company names (AIO, Cretesol Technologies, United Sol), university names, and global software brand names in standard form.
"""
    else:  # Summary & Experience Bullets Only
        lang_instructions += f"""
- Keep section headers (PROFESSIONAL SUMMARY, WORK EXPERIENCE, EDUCATION, CORE SKILLS, TOOLS, LANGUAGES) and Job Titles in standard English.
- Translate only the descriptive narrative sentences inside PROFESSIONAL SUMMARY and WORK EXPERIENCE bullet points into {target_lang}.
- Keep technical tools, methodologies, metrics, and numbers unchanged.
"""

    prompt = f"""
You are an expert European ATS Resume Writer and Recruiter specializing in the Baltic and EU tech markets.
Your task is to tailor a clean, ATS-compliant, single-column Europass CV for candidate Shafay Munir based strictly on his master background and the target Job Description.

MASTER DATA:
{PROFILES[profile_name]}

TARGET JOB DESCRIPTION:
{job_desc}

{lang_instructions}

RULES & GUIDELINES:
1. Tone & Culture: Factual, professional, metric-driven (for tech roles) or reliability-focused (for operations/services). No fluff.
2. ATS Match: Extract critical keywords from the Job Description and seamlessly integrate them into achievements and skills.
3. Truthfulness: Do not invent fake companies or fake degrees. Use the numbers and milestones provided (e.g. 50+ clients, 1.02M impressions, $17M Series A, 8-module spec).
4. FORMAT:
Output strictly structured markdown without preamble or code fencing:

# SHAFAY MUNIR
[Target Job Title matching the JD] | Vilnius, Lithuania | +370 692 28 728 | shafaymunir890@gmail.com | linkedin.com/in/shafay-munir/

# PROFESSIONAL SUMMARY
(3-4 lines targeted directly to the role, embedding key JD requirements)

# WORK EXPERIENCE
(Role, Company, Dates, Location, followed by 4-6 impact bullet points with action verbs and metrics)

# EDUCATION
(Degree, Institution, Dates)

# CORE SKILLS & EXPERTISE
(Bulleted list of exact matching technical & operational skills)

# TOOLS & PLATFORMS
(Bulleted list of software and tools relevant to the JD)

# LANGUAGES
- English (Fluent / C1-C2)
- Lithuanian (Basic / Learning)
"""
    # 1. API key ke live models dynamically fetch karein
    live_supported = []
    try:
        for m in client.models.list():
            methods = getattr(m, 'supported_generation_methods', []) or getattr(m, 'supported_actions', [])
            m_name = m.name.replace("models/", "")
            if not methods or "generateContent" in methods:
                live_supported.append(m_name)
    except Exception:
        pass

    # 2. Recommended priority order
    priority = [
        custom_model_name.strip() if custom_model_name else "",
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-3.8-flash",
        "gemini-flash-latest"
    ]
    priority.extend(live_supported)

    # 3. Known deprecated models filter out karein
    blocked = {"gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"}
    candidate_models = [m for m in dict.fromkeys(priority) if m and m not in blocked]

    last_err = None
    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt
                )
                if response and response.text:
                    return response.text
            except Exception as err:
                last_err = err
                time.sleep(1.0)
                continue

    raise last_err

def create_europass_docx(markdown_content):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.6)
        s.bottom_margin = Inches(0.6)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    COLOR_BLUE = RGBColor(0, 51, 153)
    COLOR_BODY = RGBColor(40, 40, 40)

    lines = markdown_content.strip().split("\n")
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        if line.startswith("# "):
            title = line.replace("# ", "").strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(3)
            run = p.add_run(title.upper())
            run.font.name = "Arial"
            run.font.size = Pt(11)
            run.font.bold = True
            run.font.color.rgb = COLOR_BLUE

        elif line.startswith("## "):
            subtitle = line.replace("## ", "").strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(subtitle)
            run.font.name = "Arial"
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = COLOR_BODY

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

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Hero Banner
st.markdown("""
<div class="hero-header">
    <div class="hero-title">⚡ Shafay Munir — Autonomous ATS Tailor</div>
    <div class="hero-desc">Multi-lingual reverse-chronological Europass CV generation calibrated for the Baltic & EU market.</div>
    <div class="status-chip">● ENGINE ACTIVE: MULTI-LANGUAGE ATS LOCALIZER & AUTO-HEALING</div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1.1, 1.9], gap="large")

with col1:
    st.markdown("### ⚙️ Pipeline Configuration")
    api_key = st.text_input("Gemini API Key", type="password", help="Personal AI Studio API Key")
    model_choice = st.text_input("Model Engine ID", value="gemini-2.5-flash")
    target_track = st.selectbox("Select Target Track", list(PROFILES.keys()))
    
    st.markdown("---")
    st.markdown("### 🌐 Language & Localization")
    target_language = st.selectbox("Target Output Language", LANGUAGES_AVAILABLE, index=0)
    translation_scope = st.radio(
        "Translation Scope",
        ["Entire CV", "Summary & Experience Bullets Only"],
        index=0,
        help="'Entire CV' translates headers and titles; 'Summary & Experience' preserves English technical terms and headers."
    )
    
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
        with st.status(f"⚡ Localizing & Compiling CV ({target_language})...", expanded=True) as status:
            st.write("🔍 Parsing Job Description intent and high-frequency ATS tokens...")
            time.sleep(0.3)
            st.write(f"🧠 Synthesizing Master Background for track: **{target_track}**...")
            st.write(f"🌍 Applying language localization: **{target_language}** ({translation_scope})...")
            try:
                cv_markdown = generate_tailored_cv(
                    target_track, 
                    job_desc, 
                    api_key, 
                    model_choice, 
                    target_language, 
                    translation_scope
                )
                st.write("📄 Structuring Europass Word XML document...")
                docx_file = create_europass_docx(cv_markdown)
                status.update(label="✅ Compilation Complete!", state="complete", expanded=False)
                
                clean_name = target_track.split()[0]
                lang_code = target_language.split()[0][:3].upper()
                
                st.success("✔ Document compiled successfully!")
                st.download_button(
                    label=f"📥 Download Tailored CV ({clean_name}_{lang_code}.docx)",
                    data=docx_file,
                    file_name=f"Shafay_Munir_CV_{clean_name}_{lang_code}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                with st.expander("👁️ View Live Markdown Preview"):
                    st.markdown(cv_markdown)
                    
            except Exception as e:
                status.update(label="❌ Pipeline Failed", state="error", expanded=True)
                st.error(f"Execution Error: {e}")
