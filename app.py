import os
import re
import streamlit as st
from typing import List, Dict
import pdfplumber
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
import tempfile
import base64

# ========== CONFIG ==========
try:
    # Securely fetch the API key from Streamlit's secrets
    DEEPSEEK_API_KEY = st.secrets["deepseek"]["api_key"]
except (KeyError, FileNotFoundError):
    st.error("DeepSeek API key not found. Please add it to your Streamlit secrets.")
    st.stop()

DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"


# ==========================
# Report & Infographic Structures
# ==========================
REPORT_TEMPLATES = {
"Spin-Off or Split-Up": """
Transaction Overview
ParentCo and SpinCo details
Rationale (regulatory, strategic unlock, valuation arbitrage)
Distribution terms (ratio, eligibility, tax treatment)
ParentCo Post-Spin Outlook
Strategic focus
Financial profile and valuation
SpinCo Investment Case
Business model, growth drivers
Historical and pro forma financials
Independent valuation (e.g., Sum-of-the-Parts)
Valuation Analysis
Risks and Overhangs
Forced selling, low float, governance concerns
""",
"Mergers & Acquisitions": """
Deal Summary
Parties involved, consideration (cash/stock), premium
Regulatory/antitrust/board approval status
Target Company Analysis
Valuation vs. offer
Control premium vs. peers
Buyer’s Rationale and Financing
Strategic fit
Synergies and pro forma financials
Deal financing (debt, equity)
Shareholder Vote & Antitrust Risk
Key holders' stance
Timing and likelihood of deal closure
Spread Analysis and Arbitrage Opportunity
Deal spread
IRR scenarios based on timing/risk
""",
"Bankruptcy / Distressed / Restructuring": """
Situation Summary
Cause of distress
Filing date, jurisdiction, DIP terms
Capital Structure Analysis
Pre- and post-reorg structure
Seniority waterfall
Creditor classes and recovery potential
Valuation and Recovery Scenarios
Estimated Enterprise Value
Recovery per instrument (bonds, equity, unsecured)
Reorganization Plan and Exit Timeline
Conversion to equity, rights offering, warrants
Exit multiples
Catalysts and Legal Risks
Judge approval, creditor objections, asset sales
""",
"Activist Campaign": """
Activist Background
Fund profile, history, prior campaigns
Campaign Details
Demands (board seat, spin, buyback, etc.)
Timeline of engagement
Company's Response and Governance Profile
Management alignment, shareholder defense
Scenario Analysis
Status quo vs. activist success
Proxy fight implications
Valuation Impact
NPV of potential changes (e.g., spin-off value, ROIC uplift)
""",
"Regulatory or Legal Catalyst": """
Legal/Regulatory Background
Case/issue summary
Historical legal proceedings
Outcome Scenarios
Win, loss, settlement
Timeline
Financial and Strategic Implications
Fines, product approval, license loss
Revenue/EBITDA impact
Market Reaction History (if any)
Past similar cases
""",
"Asset Sales or Carve-Outs": """
Transaction Overview
Buyer, price, structure
Valuation vs. book and peers
Strategic Impact
Focus shift, deleveraging, margin profile
Use of Proceeds
Debt repayment, dividends, buybacks, capex
Re-rating Potential
EBITDA margin uplift, return metrics
""",
"Capital Raising or Buyback Catalyst": """
Transaction Mechanics
Size, dilution, instrument type
Capital Structure Post-Deal
Leverage ratios, interest burden
Shareholder Implications
Accretion/dilution
EPS impact
Buyback Analysis (if applicable)
Repurchase pace, valuation support
"""
}

FALLBACK_META = [
    ("💼", "border-blue-600", "bg-blue-50"),
    ("🏢", "border-sky-600", "bg-sky-50"),
    ("🌐", "border-indigo-600", "bg-indigo-50"),
    ("🧩", "border-purple-600", "bg-purple-50"),
    ("📊", "border-green-600", "bg-green-50"),
    ("📈", "border-emerald-600", "bg-emerald-50"),
    ("👥", "border-yellow-600", "bg-yellow-50"),
    ("⚠️", "border-red-600", "bg-red-50"),
    ("💡", "border-pink-600", "bg-pink-50"),
    ("🧠", "border-gray-600", "bg-gray-50"),
]



# ---------------------------
# LOGO BASE64
# ---------------------------
def get_base64_logo(path="logo.png"):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

logo_base64 = get_base64_logo()

# ---------------------------
# HEADER: Smaller logo + Title below
# ---------------------------
st.markdown(f"""
    <div style="display: flex; flex-direction: column; align-items: flex-start; gap: 0.5rem; margin-bottom: 1.5rem;">
        <img src="data:image/png;base64,{logo_base64}" style="height: 36px; width: auto;" />
        
    </div>
""", unsafe_allow_html=True)

# ==========================
# Text Extractors
# ==========================

def extract_text_from_pdf(file):
    try:
        with pdfplumber.open(file) as pdf:
            return "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
    except Exception as e:
        return f"[ERROR extracting PDF: {e}]"

def extract_text_from_docx(file):
    try:
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        return f"[ERROR extracting DOCX: {e}]"

# ==========================
# Memo Generation
# ==========================

def clean_markdown(text):
    text = re.sub(r'#+\s*', '', text)
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'`{1,3}(.*?)`{1,3}', r'\1', text)
    text = re.sub(r'!\[.*?\]\(.*?\)', '', text)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'^- ', '• ', text, flags=re.MULTILINE)
    return text.strip()

def truncate_safely(text, limit=7000):
    return text[:limit]

def generate_special_situation_note(
    company_name: str,
    situation_type: str,
    uploaded_files: list,
    valuation_mode: str = None,
    parent_peers: str = "",
    spinco_peers: str = ""
):
    combined_text = ""
    for file in uploaded_files:
        if file.name.endswith(".pdf"):
            combined_text += extract_text_from_pdf(file) + "\n"
        elif file.name.endswith(".docx"):
            combined_text += extract_text_from_docx(file) + "\n"
        else:
            combined_text += f"[Unsupported file: {file.name}]\n"

    structure = REPORT_TEMPLATES.get(situation_type)
    if not structure:
        raise ValueError(f"Unsupported situation type: {situation_type}")

    # --- Conditional Valuation Section for Spin-Offs ---
    valuation_section = ""
    if situation_type == "Spin-Off or Split-Up" and valuation_mode:
        if valuation_mode == "I'll enter tickers":
            valuation_section = f"""
# Valuation Analysis
The user has provided the following peer tickers:
- ParentCo Peers: {parent_peers}
- SpinCo Peers: {spinco_peers}

Please fetch or approximate public LTM EV/EBITDA and P/E multiples for these peers.
Then apply these multiples to the extracted LTM financials of ParentCo and SpinCo to estimate standalone valuations.
Compare the sum of these to the pre-spin ParentCo's market cap to estimate the value unlock potential.
"""
        else:
            valuation_section = """
# Valuation Analysis
Based on the business descriptions of ParentCo and SpinCo, identify 3–5 appropriate public peer companies for each.
Then, estimate their LTM EV/EBITDA and P/E multiples, apply them to the extracted financials of ParentCo and SpinCo,
and compute implied valuations. Finally, compare the combined value to the pre-spin ParentCo market cap and indicate
the potential value unlock.
"""

    # --- Prompt Assembly ---
    prompt = f"""
You are an institutional investment analyst writing a professional memo on a special situation involving {company_name}.
The situation is: **{situation_type}**

Below is the internal company information extracted from various files:
\"\"\"{truncate_safely(combined_text)}\"\"\"

{valuation_section}

Using the structure below, generate a well-written investment memo. Be factual, insightful, and clear.
Structure:
{structure}
"""

    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }

    response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload)
    response.raise_for_status()
    memo = response.json()["choices"][0]["message"]["content"]

    memo = clean_markdown(memo)
    memo_dict = split_into_sections(memo, structure)

    doc = format_memo_docx(memo_dict, company_name, situation_type)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        return tmp.name


def split_into_sections(text: str, template: str) -> Dict[str, str]:
    sections = {}
    titles = [line.split('(')[0].strip() for line in template.strip().split('\n') if line.strip()]
    if not titles:
        return {"Memo": text.strip()}

    pattern = re.compile(r'^(' + '|'.join(map(re.escape, titles)) + r')\s*$', re.MULTILINE | re.IGNORECASE)
    matches = list(pattern.finditer(text))

    if not matches:
        return {"Memo": text.strip()}

    for i, match in enumerate(matches):
        title = match.group(1).strip()
        start_of_content = match.end()
        end_of_content = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start_of_content:end_of_content].strip()
        canonical_title = next((t for t in titles if t.lower() == title.lower()), title)
        sections[canonical_title] = content

    return sections

def format_memo_docx(memo_dict: dict, company_name: str, situation_type: str):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Aptos Display'
    style.font.size = Pt(11)

    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(f"{company_name} – {situation_type} Investment Memo")
    title_run.font.name = 'Aptos Display'
    title_run.font.size = Pt(20)
    title_run.bold = True
    doc.add_paragraph()

    for section_title, content in memo_dict.items():
        heading = doc.add_paragraph()
        run = heading.add_run(section_title)
        run.bold = True
        run.font.size = Pt(14)
        for para in content.strip().split('\n\n'):
            if para.strip():
                p = doc.add_paragraph(para.strip())
                p.paragraph_format.space_after = Pt(10)
                p.paragraph_format.line_spacing = 1.5
        doc.add_paragraph()

    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    
    return doc

# ==========================
# Infographic Generation
# ==========================

def extract_sections_from_docx_for_infographic(file, situation_type: str) -> Dict[str, str]:
    toc = REPORT_TEMPLATES.get(situation_type)
    if not toc:
        return {}
    
    expected_titles = {t.strip().lower() for t in toc.strip().splitlines() if t.strip()}
    doc = Document(file)
    sections = {}
    current_heading = None
    current_text = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        if text.lower() in expected_titles:
            if current_heading and current_text:
                sections[current_heading] = "\n".join(current_text).strip()
            current_heading = text
            current_text = []
        elif current_heading:
            current_text.append(text)

    if current_heading and current_text:
        sections[current_heading] = "\n".join(current_text).strip()
    
    return sections

def summarize_section_with_deepseek(section_title, section_text):
    prompt = f"""
You are an institutional research analyst preparing a financial infographic.
Summarize the section titled \"{section_title}\" into 3 to 5 concise bullet points.
Each point should be a single sentence, highlighting key insights clearly and professionally.
Section:
\"\"\"{section_text}\"\"\"
"""
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}"}
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3
    }
    response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()["choices"][0]["message"]["content"].strip()

def build_infographic_html(company_name, sections):
    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <title>{company_name} – Infographic</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        body {{ font-family: 'Inter', sans-serif; background-color: #f9fafb; color: #1f2937; }}
        .section-icon {{ font-size: 1.4rem; margin-right: 0.6rem; }}
    </style>
</head>
<body class="px-4 py-8 md:px-6 md:py-10 max-w-7xl mx-auto">
    <header class="text-center mb-12">
        <h1 class="text-3xl md:text-4xl font-bold text-gray-800 mb-2">{company_name} – Investment Memo Infographic</h1>
        <p class="text-sm text-gray-500">Generated by Aranca AI Platform</p>
    </header>
    <main class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
"""
    with st.spinner("Summarizing sections for infographic..."):
        for idx, (title, section_text) in enumerate(sections.items()):
            icon, border_class, bg_class = FALLBACK_META[idx % len(FALLBACK_META)]
            try:
                summary = summarize_section_with_deepseek(title, section_text)
                cleaned_summary = summary.replace('**', '').replace('###', '').replace('##', '').replace('#', '')
                lines = [line.lstrip("•*- ").strip() for line in cleaned_summary.split("\n") if line.strip()]
                bullet_items = "\n".join(f"            <li>{line}</li>" for line in lines)
            except Exception as e:
                bullet_items = f"<li>Error generating summary: {e}</li>"
                st.warning(f"Could not summarize section: '{title}'")

            html += f"""
        <div class="shadow-lg rounded-xl p-5 transition-transform hover:scale-[1.02] duration-300 ease-in-out border-l-4 {border_class} {bg_class}">
            <h2 class="text-lg font-semibold text-gray-800 mb-3 flex items-center">
                <span class="section-icon">{icon}</span>{title}
            </h2>
            <ul class="list-disc text-sm text-gray-700 space-y-2 pl-5 leading-relaxed">
{bullet_items}
            </ul>
        </div>
"""
    html += """
    </main>
    <footer class="text-center mt-12">
        <p class="text-xs text-gray-400">This document is for informational purposes only. Not investment advice.</p>
    </footer>
</body>
</html>
"""
    return html

# ==========================
# Streamlit App UI
# ==========================

st.set_page_config(page_title="Special Situation Memo & Infographic Generator", layout="wide")

st.title("📝 Special Situation Memo & Infographic Generator")
st.markdown("---")

st.sidebar.info("API key loaded from secrets.")

# --- Step 1: Memo Generation ---
st.header("Step 1: Generate Investment Memo")

company_name_memo = st.text_input("Enter Company Name", key="company_name_memo")
situation_type_memo = st.selectbox("Select Situation Type", options=list(REPORT_TEMPLATES.keys()), key="situation_type_memo")
use_valuation_model = False
if situation_type_memo == "Spin-Off or Split-Up":
    st.markdown("### 🔍 Valuation Module (Optional)")

    valuation_mode = st.radio(
        "Do you want to provide peer tickers for valuation, or let the model decide?",
        options=["Let AI choose peers", "I'll enter tickers"],
        key="valuation_mode"
    )

    if valuation_mode == "I'll enter tickers":
        parent_peers = st.text_input("Enter ParentCo Peer Tickers (comma-separated)", key="parent_peers")
        spinco_peers = st.text_input("Enter SpinCo Peer Tickers (comma-separated)", key="spinco_peers")
    else:
        st.info("AI will select peers using company descriptions and generate valuation logic automatically.")

    use_valuation_model = True  # activate inside memo logic
uploaded_files_memo = st.file_uploader("Upload Public Documents (PDF, DOCX)", accept_multiple_files=True, key="uploaded_files_memo")

if st.button("Generate Memo"):
    if not company_name_memo or not situation_type_memo or not uploaded_files_memo:
        st.warning("Please fill in all fields and upload at least one document.")
    else:
        with st.spinner("Generating memo... This may take a moment."):
            try:
                memo_path = generate_special_situation_note(company_name_memo, situation_type_memo, uploaded_files_memo)
                st.session_state.memo_path = memo_path
                st.session_state.company_name = company_name_memo
                st.session_state.situation_type = situation_type_memo
                
                st.success("Memo generated successfully!")
                with open(memo_path, "rb") as f:
                    st.download_button(
                        label="Download Memo (.docx)",
                        data=f,
                        file_name=f"{company_name_memo}_{situation_type_memo}_Memo.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"An error occurred during memo generation: {e}")

st.markdown("\n\n---\n\n")

# --- Step 2: Infographic Generation ---
st.header("Step 2: Generate Infographic from Memo")
st.info("After generating the memo, upload the `.docx` file below to create an infographic.")

company_name_infographic = st.text_input("Confirm Company Name", value=st.session_state.get('company_name', ''), key="company_name_infographic")
situation_type_infographic = st.selectbox("Confirm Situation Type", options=list(REPORT_TEMPLATES.keys()), index=list(REPORT_TEMPLATES.keys()).index(st.session_state.get('situation_type')) if st.session_state.get('situation_type') else 0, key="situation_type_infographic")
uploaded_memo_infographic = st.file_uploader("Upload the generated Memo (.docx)", type=["docx"], key="uploaded_memo_infographic")


if st.button("Generate Infographic"):
    if not uploaded_memo_infographic or not company_name_infographic or not situation_type_infographic:
        st.warning("Please upload the memo and confirm the company name and situation type.")
    else:
        with st.spinner("Extracting sections..."):
            try:
                sections = extract_sections_from_docx_for_infographic(uploaded_memo_infographic, situation_type_infographic)
                if not sections:
                     st.error("Could not extract any sections from the document. Please ensure the memo contains headings matching the selected situation type.")
                else:
                    st.success(f"Successfully extracted {len(sections)} sections from the memo.")
                    html_content = build_infographic_html(company_name_infographic, sections)
                    
                    st.subheader("Infographic Preview")
                    st.components.v1.html(html_content, height=800, scrolling=True)

                    st.download_button(
                        label="Download Infographic (.html)",
                        data=html_content,
                        file_name=f"{company_name_infographic}_Infographic.html",
                        mime="text/html"
                    )
            except Exception as e:
                st.error(f"An error occurred during infographic generation: {e}")