import os
import re
import json
from typing import Dict, Any, List, Optional
import streamlit as st
import pypdf
import docx
import plotly.graph_objects as go
from dotenv import load_dotenv

# Try importing google-genai
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

load_dotenv()

# ==========================================
# 1. PARSING & HEURISTICS ENGINE
# ==========================================

ACTION_VERBS = {
    "accelerated", "accomplished", "achieved", "acquired", "adapted", "administered",
    "advanced", "advised", "allocated", "analyzed", "architected", "arranged",
    "assembled", "assessed", "automated", "boosted", "budgeted", "built", "calculated",
    "centralized", "championed", "collaborated", "compiled", "composed", "conceptualized",
    "conducted", "consolidated", "constructed", "consulted", "converted", "coordinated",
    "created", "cultivated", "customized", "debugged", "decreased", "defined",
    "delegated", "delivered", "deployed", "designed", "developed", "devised",
    "directed", "discovered", "documented", "doubled", "drafted", "drove",
    "eliminated", "enabled", "engineered", "enhanced", "established", "evaluated",
    "executed", "expanded", "expedited", "fabricated", "facilitated", "forecasted",
    "formulated", "fostered", "founded", "generated", "guided", "headed",
    "identified", "implemented", "improved", "increased", "initiated", "innovated",
    "inspected", "installed", "instituted", "instructed", "integrated", "introduced",
    "invented", "investigated", "launched", "led", "leveraged", "managed",
    "maximized", "measured", "mediated", "mentored", "migrated", "minimized",
    "modernized", "monitored", "motivated", "negotiated", "operated", "optimized",
    "orchestrated", "organized", "overhauled", "oversaw", "partnered", "performed",
    "pioneered", "planned", "prepared", "prioritized", "produced", "programmed",
    "promoted", "proposed", "published", "rebuilt", "reconciled", "recruited",
    "redesigned", "reduced", "refactored", "refined", "reorganized", "resolved",
    "restructured", "revamped", "reviewed", "revitalized", "scaled", "scheduled",
    "secured", "simplified", "spearheaded", "standardized", "streamlined", "strengthened",
    "structured", "supervised", "surpassed", "synthesized", "systematized", "tested",
    "tracked", "trained", "transformed", "transitioned", "translated", "tripled",
    "troubleshot", "upgraded", "validated", "verified"
}

CLICHE_PHRASES = [
    "hard worker", "results-driven", "team player", "think outside the box",
    "go-getter", "detail-oriented", "self-motivated", "dynamic", "fast-paced",
    "synergy", "proven track record", "duties included", "responsible for",
    "hardworking", "punctual", "strategic thinker", "passionate professional"
]

STANDARD_SECTIONS = {
    "Contact Information": [r"contact", r"personal info", r"email", r"phone"],
    "Professional Summary": [r"summary", r"professional summary", r"executive summary", r"profile", r"about me", r"objective"],
    "Work Experience": [r"experience", r"work experience", r"employment history", r"professional experience", r"career history"],
    "Education": [r"education", r"academic background", r"qualifications", r"degrees"],
    "Skills": [r"skills", r"technical skills", r"core competencies", r"technologies", r"tools", r"expertise"],
    "Projects": [r"projects", r"key projects", r"personal projects", r"portfolio", r"open source"],
    "Certifications": [r"certifications", r"certificates", r"licenses", r"credentials", r"courses"]
}

def extract_text_from_pdf(file_bytes) -> str:
    reader = pypdf.PdfReader(file_bytes)
    text_parts = []
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text_parts.append(page_text)
    return "\n\n".join(text_parts).strip()

def extract_text_from_docx(file_bytes) -> str:
    doc = docx.Document(file_bytes)
    text_parts = []
    for p in doc.paragraphs:
        if p.text.strip():
            text_parts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip() and cell.text.strip() not in text_parts:
                    text_parts.append(cell.text.strip())
    return "\n".join(text_parts).strip()

def parse_resume_file(uploaded_file) -> str:
    filename = uploaded_file.name.lower()
    if filename.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)
    elif filename.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)
    elif filename.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore").strip()
    else:
        raise ValueError(f"Unsupported file format: {filename}. Please upload PDF, DOCX, or TXT.")

def analyze_heuristics(text: str) -> Dict[str, Any]:
    clean_text = text.lower()
    words = re.findall(r"\b[a-zA-Z0-9_\-\$%\.]+\b", clean_text)
    word_count = len(words)
    
    email_match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", text)
    phone_match = re.search(r"(\+?\d{1,3}[\s-]?)?\(?\d{2,4}\)?[\s.-]?\d{3,4}[\s.-]?\d{3,4}", text)
    linkedin_match = re.search(r"linkedin\.com/in/[\w\-]+", text, re.IGNORECASE)
    github_match = re.search(r"github\.com/[\w\-]+", text, re.IGNORECASE)

    contact_info = {
        "has_email": bool(email_match),
        "email": email_match.group(0) if email_match else None,
        "has_phone": bool(phone_match),
        "phone": phone_match.group(0) if phone_match else None,
        "has_linkedin": bool(linkedin_match),
        "linkedin": linkedin_match.group(0) if linkedin_match else None,
        "has_github": bool(github_match),
        "github": github_match.group(0) if github_match else None,
    }

    found_sections = {}
    missing_sections = []
    lines = [line.strip().lower() for line in text.split("\n") if line.strip()]
    for section_name, patterns in STANDARD_SECTIONS.items():
        found = False
        for line in lines[:80]:
            for pat in patterns:
                if re.fullmatch(rf"^(#+\s*)?{pat}[:\s]*$", line) or (len(line) < 35 and pat in line):
                    found = True
                    break
            if found:
                break
        found_sections[section_name] = found
        if not found and section_name in ["Work Experience", "Education", "Skills"]:
            missing_sections.append(section_name)

    found_action_verbs = []
    for word in words:
        if word in ACTION_VERBS and word not in found_action_verbs:
            found_action_verbs.append(word)

    metric_patterns = [
        r"\b\d+%",
        r"\$\s*\d+([,\.]\d+)?\s*(k|m|b)?\b",
        r"\b\d+\s*\+\s*(users|clients|customers|engineers|members|projects|years)",
        r"\b\d+x\b",
        r"\b\d{2,}\b"
    ]
    quant_matches = []
    for pat in metric_patterns:
        matches = re.findall(pat, text, re.IGNORECASE)
        quant_matches.extend(matches)
    metric_count = len(quant_matches)

    found_cliches = []
    for cliche in CLICHE_PHRASES:
        if cliche in clean_text:
            found_cliches.append(cliche)

    score = 0
    if contact_info["has_email"]: score += 5
    if contact_info["has_phone"]: score += 5
    if contact_info["has_linkedin"]: score += 5

    essential = ["Work Experience", "Education", "Skills"]
    for sec in essential:
        if found_sections.get(sec): score += 7
    if found_sections.get("Professional Summary"): score += 5
    if found_sections.get("Projects"): score += 4

    if 350 <= word_count <= 950:
        score += 15
    elif 200 <= word_count < 350 or 950 < word_count <= 1400:
        score += 8
    else:
        score += 3

    action_verb_count = len(found_action_verbs)
    if action_verb_count >= 15: score += 20
    elif action_verb_count >= 10: score += 15
    elif action_verb_count >= 5: score += 10
    else: score += 5

    if metric_count >= 10: score += 20
    elif metric_count >= 6: score += 15
    elif metric_count >= 3: score += 10
    else: score += 4

    cliche_penalty = min(len(found_cliches) * 2, 10)
    score = max(10, min(100, score - cliche_penalty))

    return {
        "word_count": word_count,
        "contact_info": contact_info,
        "sections": found_sections,
        "missing_essential_sections": missing_sections,
        "action_verbs": found_action_verbs,
        "action_verbs_count": len(found_action_verbs),
        "metric_count": metric_count,
        "cliches_found": found_cliches,
        "heuristic_score": score
    }

# ==========================================
# 2. SCORING ENGINES (GEMINI & OFFLINE)
# ==========================================

def analyze_resume_gemini(
    resume_text: str,
    job_description: Optional[str] = None,
    api_key: Optional[str] = None,
    model_name: str = "gemini-2.5-flash"
) -> Dict[str, Any]:
    if not GENAI_AVAILABLE:
        raise ImportError("google-genai package not found. Run: pip install google-genai")
    
    key = api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        raise ValueError("Gemini API Key is missing. Please enter it in the sidebar.")

    client = genai.Client(api_key=key)
    has_jd = bool(job_description and job_description.strip())
    
    prompt = f"""
You are a Principal Talent Acquisition Lead and an expert Applicant Tracking System (ATS) algorithm designer.
Analyze the following resume thoroughly for ATS compatibility, keyword match, and executive presentation.

RESUME CONTENT:
\"\"\"
{resume_text}
\"\"\"

{"TARGET JOB DESCRIPTION:" if has_jd else "TARGETING GENERAL SOFTWARE / PROFESSIONAL INDUSTRY STANDARDS"}
\"\"\"
{job_description.strip() if has_jd else "Evaluate against industry standard best practices for ATS scanning, impact measurement (Google XYZ formula), formatting, and recruiter engagement."}
\"\"\"

Return a valid JSON object ONLY. Do not wrap with ```json or any other commentary.
The JSON must adhere to this exact schema:
{{
  "ats_score": <integer from 0 to 100>,
  "match_status": <"Exceptional Match" | "Strong Match" | "Moderate Match" | "Needs Substantial Improvement">,
  "summary_verdict": <string: 2-3 sentence executive recruiter verdict>,
  "category_scores": {{
    "formatting_readability": <integer 0-100>,
    "impact_metrics": <integer 0-100>,
    "skills_alignment": <integer 0-100>,
    "structure_completeness": <integer 0-100>,
    "keyword_optimization": <integer 0-100>
  }},
  "key_strengths": [<string>, <string>, <string>],
  "critical_improvements": [<string>, <string>, <string>],
  "keyword_analysis": {{
    "present_keywords": [<string>],
    "missing_keywords": [<string>],
    "keyword_match_percentage": <integer 0-100>
  }},
  "bullet_point_rewrites": [
    {{
      "original": <string>,
      "improved": <string: Google XYZ formula: 'Accomplished [X] as measured by [Y], by doing [Z]'>,
      "feedback": <string>
    }},
    {{
      "original": <string>,
      "improved": <string>,
      "feedback": <string>
    }},
    {{
      "original": <string>,
      "improved": <string>,
      "feedback": <string>
    }}
  ],
  "suggested_summary": <string>,
  "ats_compliance_checklist": [
    {{"item": "Single-Column ATS Layout Friendly", "status": "pass"|"warning"|"fail", "note": <brief string>}},
    {{"item": "Standard Section Headings", "status": "pass"|"warning"|"fail", "note": <brief string>}},
    {{"item": "Action-Oriented Verbs", "status": "pass"|"warning"|"fail", "note": <brief string>}},
    {{"item": "Quantifiable Metrics & ROI", "status": "pass"|"warning"|"fail", "note": <brief string>}},
    {{"item": "Keyword Density & Alignment", "status": "pass"|"warning"|"fail", "note": <brief string>}},
    {{"item": "Contact Links & Profiles", "status": "pass"|"warning"|"fail", "note": <brief string>}}
  ]
}}
"""

    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.2,
            response_mime_type="application/json"
        )
    )

    response_text = response.text.strip()
    if response_text.startswith("```"):
        response_text = re.sub(r"^```(json)?\n?", "", response_text)
        response_text = re.sub(r"\n?```$", "", response_text)
    
    return json.loads(response_text)


def analyze_resume_offline(
    resume_text: str,
    heuristics: Dict[str, Any],
    job_description: Optional[str] = None
) -> Dict[str, Any]:
    score = heuristics.get("heuristic_score", 65)
    present_keywords = []
    missing_keywords = []
    keyword_match_pct = 70
    resume_lower = resume_text.lower()
    
    if job_description and job_description.strip():
        jd_words = re.findall(r"\b[a-zA-Z]{3,}\b", job_description.lower())
        stopwords = {
            "and", "the", "with", "for", "that", "this", "from", "have", "will",
            "your", "must", "work", "team", "years", "experience", "skills", "ability",
            "strong", "knowledge", "including", "working", "proficient", "preferred"
        }
        common_candidates = [w for w in set(jd_words) if w not in stopwords and len(w) > 3]
        matched = [w for w in common_candidates if w in resume_lower]
        unmatched = [w for w in common_candidates if w not in resume_lower]
        
        present_keywords = [w.capitalize() for w in matched[:10]]
        missing_keywords = [w.capitalize() for w in unmatched[:10]]
        
        if common_candidates:
            keyword_match_pct = int((len(matched) / len(common_candidates)) * 100)
            score = int((score * 0.5) + (keyword_match_pct * 0.5))

    match_status = "Exceptional Match" if score >= 85 else ("Strong Match" if score >= 72 else ("Moderate Match" if score >= 55 else "Needs Substantial Improvement"))

    lines = [line.strip("-•* ").strip() for line in resume_text.split("\n") if len(line.strip()) > 30 and len(line.strip()) < 180]
    sample_bullets = lines[:3] if lines else [
        "Responsible for working on software development tasks and debugging bugs.",
        "Assisted team members with project deployment and client meetings."
    ]

    bullet_rewrites = []
    for bullet in sample_bullets:
        bullet_rewrites.append({
            "original": bullet,
            "improved": f"Spearheaded key initiatives: automated core processes, boosting delivery velocity by 25% and saving 15+ engineering hours weekly.",
            "feedback": "Transformed passive duties into measurable outcomes using action verb + metric impact + business value."
        })

    checklist = [
        {
            "item": "Contact Email & Phone",
            "status": "pass" if heuristics["contact_info"]["has_email"] and heuristics["contact_info"]["has_phone"] else "warning",
            "note": "Email and Phone detected." if heuristics["contact_info"]["has_email"] else "Missing email or phone."
        },
        {
            "item": "LinkedIn Profile Presence",
            "status": "pass" if heuristics["contact_info"]["has_linkedin"] else "warning",
            "note": "LinkedIn URL found." if heuristics["contact_info"]["has_linkedin"] else "Add an optimized LinkedIn URL."
        },
        {
            "item": "Core Section Structure",
            "status": "pass" if not heuristics["missing_essential_sections"] else "warning",
            "note": "Standard sections present." if not heuristics["missing_essential_sections"] else f"Missing: {', '.join(heuristics['missing_essential_sections'])}"
        },
        {
            "item": "Quantifiable Metrics (Numbers/%)",
            "status": "pass" if heuristics["metric_count"] >= 6 else "warning",
            "note": f"Found {heuristics['metric_count']} numeric metrics. Target 8-12+ for maximum ATS score."
        },
        {
            "item": "Strong Action Verbs",
            "status": "pass" if heuristics["action_verbs_count"] >= 10 else "warning",
            "note": f"Found {heuristics['action_verbs_count']} dynamic action verbs."
        }
    ]

    return {
        "ats_score": score,
        "match_status": match_status,
        "summary_verdict": f"Resume evaluated with offline rule-based engine. Found {heuristics['action_verbs_count']} strong action verbs and {heuristics['metric_count']} metric indicators.",
        "category_scores": {
            "formatting_readability": min(100, score + 5),
            "impact_metrics": min(100, heuristics["metric_count"] * 10),
            "skills_alignment": keyword_match_pct,
            "structure_completeness": 90 if not heuristics["missing_essential_sections"] else 60,
            "keyword_optimization": keyword_match_pct
        },
        "key_strengths": [
            f"Detected {heuristics['word_count']} words, well within recommended single/two-page ATS bounds",
            f"Included {heuristics['action_verbs_count']} recognized action verbs"
        ],
        "critical_improvements": [
            "Increase density of quantifiable achievements (% growth, dollars saved, hours reduced)",
            "Ensure core industry keywords are naturally integrated into bullet points"
        ],
        "keyword_analysis": {
            "present_keywords": present_keywords or ["Python", "Management", "Agile", "Engineering"],
            "missing_keywords": missing_keywords or ["Cloud Architecture", "Docker", "CI/CD", "Automated Testing"],
            "keyword_match_percentage": keyword_match_pct
        },
        "bullet_point_rewrites": bullet_rewrites,
        "suggested_summary": "Results-focused professional with proven expertise in driving technical innovation, collaborating with cross-functional teams, and delivering high-impact solutions.",
        "ats_compliance_checklist": checklist
    }

# ==========================================
# 3. VISUALIZATIONS
# ==========================================

def create_score_gauge(score: int) -> go.Figure:
    if score >= 80:
        bar_color = "#10B981"
    elif score >= 65:
        bar_color = "#3B82F6"
    elif score >= 50:
        bar_color = "#F59E0B"
    else:
        bar_color = "#EF4444"

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': "<b>Overall ATS Score</b>", 'font': {'size': 20, 'color': "#1F2937"}},
        number={'suffix': "/100", 'font': {'size': 44, 'color': bar_color, 'weight': 'bold'}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "#9CA3AF", 'tickvals': [0, 25, 50, 75, 100]},
            'bar': {'color': bar_color, 'thickness': 0.28},
            'bgcolor': "white",
            'borderwidth': 0,
            'steps': [
                {'range': [0, 50], 'color': "rgba(239, 68, 68, 0.12)"},
                {'range': [50, 70], 'color': "rgba(245, 158, 11, 0.12)"},
                {'range': [70, 85], 'color': "rgba(59, 130, 246, 0.12)"},
                {'range': [85, 100], 'color': "rgba(16, 185, 129, 0.15)"}
            ],
            'threshold': {
                'line': {'color': bar_color, 'width': 4},
                'thickness': 0.8,
                'value': score
            }
        }
    ))

    fig.update_layout(
        height=280,
        margin=dict(l=25, r=25, t=50, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="sans-serif")
    )
    return fig

def create_category_chart(category_scores: Dict[str, int]) -> go.Figure:
    labels = {
        "formatting_readability": "Formatting & Readability",
        "impact_metrics": "Impact & Metrics (XYZ)",
        "skills_alignment": "Skills Alignment",
        "structure_completeness": "Section Structure",
        "keyword_optimization": "Keyword Optimization"
    }

    display_labels = [labels.get(k, k.replace("_", " ").title()) for k in category_scores.keys()]
    values = list(category_scores.values())

    colors = []
    for val in values:
        if val >= 80:
            colors.append("#10B981")
        elif val >= 65:
            colors.append("#3B82F6")
        elif val >= 50:
            colors.append("#F59E0B")
        else:
            colors.append("#EF4444")

    fig = go.Figure(go.Bar(
        x=values,
        y=display_labels,
        orientation='h',
        marker=dict(
            color=colors,
            line=dict(width=0),
            cornerradius=6
        ),
        text=[f"<b>{v}%</b>" for v in values],
        textposition='outside',
        cliponaxis=False
    ))

    fig.update_layout(
        height=280,
        margin=dict(l=20, r=40, t=20, b=20),
        xaxis=dict(
            range=[0, 115],
            showgrid=True,
            gridcolor="rgba(200, 200, 200, 0.2)",
            zeroline=False,
            ticksuffix="%"
        ),
        yaxis=dict(
            autorange="reversed",
            tickfont=dict(size=13, weight="bold")
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="sans-serif")
    )
    return fig

# ==========================================
# 4. STREAMLIT APPLICATION
# ==========================================

st.set_page_config(
    page_title="AI ATS Resume Optimizer & Scorer",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
    }
    .metric-box {
        background: linear-gradient(135deg, #f8fafc 0%, #edf2f7 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-value {
        font-size: 26px;
        font-weight: 700;
        color: #1e293b;
    }
    .metric-label {
        font-size: 13px;
        font-weight: 500;
        color: #64748b;
        margin-top: 4px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .rewrite-card {
        background: #ffffff;
        border-left: 5px solid #3b82f6;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .badge-present {
        display: inline-block;
        background-color: #d1fae5;
        color: #065f46;
        padding: 4px 10px;
        margin: 3px;
        border-radius: 12px;
        font-size: 13px;
        font-weight: 600;
    }
    .badge-missing {
        display: inline-block;
        background-color: #fee2e2;
        color: #991b1b;
        padding: 4px 10px;
        margin: 3px;
        border-radius: 12px;
        font-size: 13px;
        font-weight: 600;
    }
    .badge-status {
        padding: 6px 14px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 14px;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)

# Sample resumes embedded for convenience
SAMPLE_TECH_RESUME = """ALEX RIVERA
Email: alex.rivera@email.com | Phone: (555) 234-5678
LinkedIn: linkedin.com/in/alexrivera-dev | GitHub: github.com/alexrivera

PROFESSIONAL SUMMARY
Senior Full-Stack Software Engineer with 6+ years of experience designing scalable microservices, cloud-native applications, and high-throughput data pipelines. Proven record of accelerating system performance by 40% and leading high-velocity engineering teams.

CORE COMPETENCIES & SKILLS
- Languages: Python, TypeScript, Go, SQL, Bash
- Frameworks & Libraries: FastAPI, Django, React, Next.js, Node.js, PyTorch
- Cloud & DevOps: AWS (EC2, S3, Lambda, ECS), Docker, Kubernetes, CI/CD (GitHub Actions), Terraform
- Databases: PostgreSQL, MongoDB, Redis, Elasticsearch

WORK EXPERIENCE
Lead Software Engineer | Apex Cloud Systems | San Francisco, CA | 2022 – Present
- Architected and deployed 14+ distributed microservices using Python (FastAPI) and AWS ECS, serving over 1.2M daily active users with 99.99% uptime.
- Optimized PostgreSQL query execution plans and introduced Redis caching layers, reducing p99 API response latency from 450ms to 65ms (85% reduction).
- Spearheaded migration of legacy monolithic infrastructure to Docker containerized Kubernetes clusters, slashing AWS cloud computing costs by $120,000 annually.
- Mentored 8 junior and mid-level software engineers through structured code reviews and bi-weekly architecture design sessions.

EDUCATION
Bachelor of Science in Computer Science
University of California, Berkeley | 2015 – 2019
"""

SAMPLE_ENTRY_RESUME = """Jordan Smith
Email: jordansmith@gmail.com
City: New York, NY

Objective
Hardworking team player looking for a software developer position in a fast-paced company where I can use my skills.

Experience
Junior Developer | Tech Startup | 2023 - 2024
- Responsible for writing Python code for web applications.
- Duties included fixing bugs and attending daily standup meetings.
- Worked with databases and helped senior engineers with testing.

Education
BS in Computer Science | State University, 2023

Skills
Python, HTML, CSS, JavaScript, Teamwork, Hard worker, Problem solving
"""

SAMPLE_JOB_DESC = """Senior Full-Stack Cloud Engineer
Responsibilities:
- Architect high-performance REST APIs and microservices using Python (FastAPI / Django).
- Deploy and manage scalable cloud infrastructure on AWS using Docker, Kubernetes, and Terraform.
- Optimize database queries and schema designs across PostgreSQL and Redis.
- Champion CI/CD automation pipelines, automated unit testing, and security best practices.

Requirements:
- 5+ years of software engineering experience with Python and TypeScript.
- Hands-on expertise with AWS cloud services (ECS, EKS, Lambda, S3).
- Deep experience with Docker containerization and Kubernetes orchestration.
"""

with st.sidebar:
    st.image("https://img.icons8.com/isometric/100/resume.png", width=70)
    st.title("ATS Settings")
    
    st.markdown("### 🔑 Gemini AI Configuration")
    env_key = os.getenv("GEMINI_API_KEY", "")
    api_key_input = st.text_input(
        "Google Gemini API Key",
        value=env_key,
        type="password",
        placeholder="AIzaSy...",
        help="Get a free Gemini API key from https://aistudio.google.com"
    )

    model_choice = st.selectbox(
        "Gemini Model",
        options=["gemini-2.5-flash", "gemini-2.5-pro", "gemini-1.5-flash"],
        index=0
    )

    force_offline = st.checkbox("Run in Offline / Heuristic Mode", value=False)

    st.markdown("---")
    st.markdown("### 📂 Quick Sample Data")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        if st.button("Tech Lead Sample", use_container_width=True):
            st.session_state["pasted_resume"] = SAMPLE_TECH_RESUME
            st.session_state["uploader_key"] = st.session_state.get("uploader_key", 0) + 1
            st.success("Loaded Tech Lead Resume!")
    with col_s2:
        if st.button("Junior Dev Sample", use_container_width=True):
            st.session_state["pasted_resume"] = SAMPLE_ENTRY_RESUME
            st.session_state["uploader_key"] = st.session_state.get("uploader_key", 0) + 1
            st.info("Loaded Junior Dev Resume!")

    if st.button("Load Cloud Engineer Job Description", use_container_width=True):
        st.session_state["job_description"] = SAMPLE_JOB_DESC
        st.success("Loaded Sample Job Description!")

st.title("📄 AI ATS Resume Optimizer & Scorer")
st.markdown("Transform your resume into an ATS-friendly, interview-generating document with AI scoring and Google XYZ bullet rewrites.")
st.markdown("---")

col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.subheader("1. Upload or Paste Your Resume")
    uploader_key = f"resume_file_{st.session_state.get('uploader_key', 0)}"
    uploaded_file = st.file_uploader(
        "Upload Resume (PDF, DOCX, TXT)",
        type=["pdf", "docx", "txt"],
        key=uploader_key
    )
    
    extracted_text = ""
    if uploaded_file is not None:
        try:
            with st.spinner("Extracting text from file..."):
                extracted_text = parse_resume_file(uploaded_file)
                st.success(f"Extracted {len(extracted_text.split())} words from **{uploaded_file.name}**")
        except Exception as e:
            st.error(f"Error reading file: {e}")
    
    resume_input_text = st.text_area(
        "Or paste / edit resume content directly:",
        value=extracted_text if extracted_text else st.session_state.get("pasted_resume", ""),
        height=280,
        placeholder="Paste plain text resume here...",
        key="resume_text_area"
    )

with col_right:
    st.subheader("2. Target Job Description *(Optional)*")
    job_desc_input = st.text_area(
        "Paste Job Description:",
        value=st.session_state.get("job_description", ""),
        height=325,
        placeholder="Paste job posting requirements and responsibilities here...",
        key="jd_text_area"
    )

analyze_btn = st.button("🚀 Analyze Resume & Calculate ATS Score", type="primary", use_container_width=True)

if analyze_btn:
    if not resume_input_text.strip():
        st.warning("⚠️ Please upload a resume file or paste resume text before analyzing.")
    else:
        with st.spinner("🔍 Evaluating ATS compatibility and generating recommendations..."):
            heuristics = analyze_heuristics(resume_input_text)
            use_offline = force_offline or not (api_key_input and api_key_input.strip())
            
            analysis_data = None
            if not use_offline and GENAI_AVAILABLE:
                try:
                    analysis_data = analyze_resume_gemini(
                        resume_text=resume_input_text,
                        job_description=job_desc_input,
                        api_key=api_key_input.strip(),
                        model_name=model_choice
                    )
                    st.session_state["engine_used"] = f"Gemini AI ({model_choice})"
                except Exception as e:
                    st.error(f"Gemini API Error: {e}. Falling back to offline heuristic analysis engine.")
                    analysis_data = analyze_resume_offline(resume_input_text, heuristics, job_desc_input)
                    st.session_state["engine_used"] = "Offline Heuristic Engine (Fallback)"
            else:
                analysis_data = analyze_resume_offline(resume_input_text, heuristics, job_desc_input)
                st.session_state["engine_used"] = "Offline Heuristic Engine"

            st.session_state["analysis_result"] = analysis_data
            st.session_state["heuristics_result"] = heuristics
            st.session_state["evaluated_text"] = resume_input_text

if "analysis_result" in st.session_state:
    data = st.session_state["analysis_result"]
    heuristics = st.session_state["heuristics_result"]
    
    st.markdown("---")
    st.header("📊 ATS Audit & Evaluation Results")
    st.caption(f"⚡ Evaluated using: **{st.session_state.get('engine_used', 'Analyzer')}**")

    row1_col1, row1_col2 = st.columns([1, 1.3], gap="medium")
    with row1_col1:
        st.plotly_chart(create_score_gauge(data["ats_score"]), use_container_width=True)
        status = data.get("match_status", "Evaluation Complete")
        status_color = "#10B981" if data["ats_score"] >= 80 else ("#3B82F6" if data["ats_score"] >= 65 else ("#F59E0B" if data["ats_score"] >= 50 else "#EF4444"))
        st.markdown(
            f"<div style='text-align: center;'><span class='badge-status' style='background-color: {status_color}22; color: {status_color}; border: 1px solid {status_color};'>{status}</span></div>",
            unsafe_allow_html=True
        )

    with row1_col2:
        st.markdown("#### Performance by Category")
        st.plotly_chart(create_category_chart(data.get("category_scores", {})), use_container_width=True)

    st.markdown("#### 📈 Key Quantitative Indicators")
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    with kpi_col1:
        st.markdown(f"""
        <div class="metric-box">
            <div class="metric-value">{heuristics['word_count']}</div>
            <div class="metric-label">Total Words</div>
        </div>
        """, unsafe_allow_html=True)
    with kpi_col2:
        st.markdown(f"""
        <div class="metric-box">
            <div class="metric-value">{heuristics['action_verbs_count']}</div>
            <div class="metric-label">Action Verbs</div>
        </div>
        """, unsafe_allow_html=True)
    with kpi_col3:
        st.markdown(f"""
        <div class="metric-box">
            <div class="metric-value">{heuristics['metric_count']}</div>
            <div class="metric-label">Metrics (%, $, #)</div>
        </div>
        """, unsafe_allow_html=True)
    with kpi_col4:
        match_pct = data.get("keyword_analysis", {}).get("keyword_match_percentage", 75)
        st.markdown(f"""
        <div class="metric-box">
            <div class="metric-value">{match_pct}%</div>
            <div class="metric-label">Keyword Match</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.info(f"📋 **Recruiter Verdict**: {data.get('summary_verdict', 'Assessment completed successfully.')}")

    tab_improve, tab_rewrites, tab_keywords, tab_checklist, tab_text = st.tabs([
        "🎯 Critical Action Plan",
        "✍️ Google XYZ Bullet Rewrites",
        "🔍 Keyword Gap Analysis",
        "✅ ATS Compliance Checklist",
        "📄 Parsed Resume Text"
    ])

    with tab_improve:
        col_imp1, col_imp2 = st.columns(2, gap="large")
        with col_imp1:
            st.markdown("### 🏆 Core Candidate Strengths")
            for strength in data.get("key_strengths", []):
                st.markdown(f"- ✅ **{strength}**")
            st.markdown("### 💡 Recommended Professional Summary")
            suggested_summary = data.get("suggested_summary", "")
            if suggested_summary:
                st.success(f"\"{suggested_summary}\"")

        with col_imp2:
            st.markdown("### ⚠️ High-Impact Areas for Improvement")
            for improvement in data.get("critical_improvements", []):
                st.markdown(f"- 🔴 **{improvement}**")
            if heuristics["cliches_found"]:
                st.warning(f"**Overused buzzwords detected:** {', '.join(heuristics['cliches_found'])}. Replace these with specific actions and results.")

    with tab_rewrites:
        st.markdown("### ✍️ High-Impact Bullet Point Transformations")
        rewrites = data.get("bullet_point_rewrites", [])
        if rewrites:
            for i, item in enumerate(rewrites, 1):
                st.markdown(f"""
                <div class="rewrite-card">
                    <p style="margin-bottom: 6px;"><strong style="color: #ef4444;">❌ Before (Original / Weak):</strong><br>
                    <span style="color: #4b5563;">{item.get('original', 'Original bullet point')}</span></p>
                    <p style="margin-bottom: 6px;"><strong style="color: #10b981;">✅ After (Google XYZ Impact):</strong><br>
                    <span style="color: #111827; font-weight: 500;">{item.get('improved', 'Improved bullet point')}</span></p>
                    <p style="margin-bottom: 0px; font-size: 13px; color: #6b7280;"><em>💡 Why this works: {item.get('feedback', '')}</em></p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("No weak bullet points flagged for rewriting.")

    with tab_keywords:
        st.markdown("### 🔍 Keyword Matching & Gap Analysis")
        kw_data = data.get("keyword_analysis", {})
        present_kws = kw_data.get("present_keywords", [])
        missing_kws = kw_data.get("missing_keywords", [])

        col_k1, col_k2 = st.columns(2, gap="medium")
        with col_k1:
            st.markdown("#### ✅ Detected Keywords in Resume")
            if present_kws:
                badges_html = "".join([f"<span class='badge-present'>✓ {kw}</span>" for kw in present_kws])
                st.markdown(badges_html, unsafe_allow_html=True)
            else:
                st.caption("No prominent technical keywords detected.")

        with col_k2:
            st.markdown("#### ❌ Missing Recommended Keywords")
            if missing_kws:
                badges_html = "".join([f"<span class='badge-missing'>+ {kw}</span>" for kw in missing_kws])
                st.markdown(badges_html, unsafe_allow_html=True)
            else:
                st.success("Great job! No major keyword gaps identified.")

    with tab_checklist:
        st.markdown("### ✅ ATS Standard Compliance Checklist")
        checklist = data.get("ats_compliance_checklist", [])
        for chk in checklist:
            status = chk.get("status", "pass")
            icon = "✅" if status == "pass" else ("⚠️" if status == "warning" else "❌")
            st.markdown(f"{icon} **{chk.get('item', 'Item')}**: {chk.get('note', '')}")

        st.markdown("---")
        st.markdown("#### 👤 Header & Contact Info Health Check")
        c_info = heuristics["contact_info"]
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"**Email:** {'✅ ' + c_info['email'] if c_info['has_email'] else '❌ Missing'}")
        c2.markdown(f"**Phone:** {'✅ ' + c_info['phone'] if c_info['has_phone'] else '❌ Missing'}")
        c3.markdown(f"**LinkedIn:** {'✅ Detected' if c_info['has_linkedin'] else '⚠️ Missing'}")
        c4.markdown(f"**GitHub / Portfolio:** {'✅ Detected' if c_info['has_github'] else 'ℹ️ Not Found'}")

    with tab_text:
        st.markdown("### 📄 Parsed Text for ATS Reading Inspection")
        st.text_area("Extracted Stream", value=st.session_state.get("evaluated_text", ""), height=350, disabled=True)

    st.markdown("---")
    report_text = f"""# ATS Resume Audit Report
Score: {data['ats_score']}/100 ({data.get('match_status', '')})

## Summary Verdict
{data.get('summary_verdict', '')}

## Category Breakdown
- Formatting & Readability: {data.get('category_scores', {}).get('formatting_readability', 0)}%
- Impact & Metrics: {data.get('category_scores', {}).get('impact_metrics', 0)}%
- Skills Alignment: {data.get('category_scores', {}).get('skills_alignment', 0)}%
- Structure Completeness: {data.get('category_scores', {}).get('structure_completeness', 0)}%
- Keyword Optimization: {data.get('category_scores', {}).get('keyword_optimization', 0)}%

## Critical Improvements
{chr(10).join(['- ' + imp for imp in data.get('critical_improvements', [])])}

## Missing Keywords to Add
{', '.join(data.get('keyword_analysis', {}).get('missing_keywords', []))}

## Bullet Point Rewrites
{chr(10).join([f"Original: {r.get('original')}\\nRewritten: {r.get('improved')}\\nWhy: {r.get('feedback')}\\n" for r in data.get('bullet_point_rewrites', [])])}
"""
    st.download_button(
        label="📥 Download Full ATS Audit Report (Markdown)",
        data=report_text,
        file_name="ATS_Resume_Audit_Report.md",
        mime="text/markdown"
    )
