import hashlib
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch
from pptx import Presentation
from pptx.util import Pt

st.set_page_config(
    page_title="GraphPilot AI | Document Intelligence",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
:root{--bg:#0B1020;--panel:#151D35;--panel2:#192440;--line:rgba(154,170,210,.16);--muted:#AAB7D3;--text:#F4F6FF;--purple:#795CFF;--blue:#168BFF}
html,body,[data-testid="stAppViewContainer"]{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif}
[data-testid="stHeader"]{background:rgba(11,16,32,.88)}
[data-testid="stSidebar"]{background:#0E162B;border-right:1px solid var(--line)}
[data-testid="stSidebar"] *{color:var(--text)}
.block-container{padding-top:1.6rem;padding-bottom:3rem;max-width:1500px}
h1,h2,h3{font-family:'Space Grotesk',sans-serif!important;color:var(--text)!important;letter-spacing:-.035em}
p,label,.stMarkdown{color:var(--text)}
.eyebrow{color:#B6A9FF;font-size:11px;font-weight:700;letter-spacing:1.8px;text-transform:uppercase}
.brand{display:flex;align-items:center;gap:11px;padding:8px 4px 24px}
.brand-mark{width:38px;height:38px;border-radius:13px;display:grid;place-items:center;background:linear-gradient(135deg,var(--purple),var(--blue));font-size:20px}
.brand-name{font:700 19px 'Space Grotesk'}
.brand-sub{color:var(--muted);font-size:10px;letter-spacing:1.7px;text-transform:uppercase}
.hero{padding:25px 28px;border:1px solid var(--line);border-radius:20px;background:radial-gradient(ellipse at 88% 0%,#5b3df533 0%,transparent 42%),linear-gradient(120deg,#172342,#111a31);margin-bottom:20px}
.hero h1{font-size:32px;margin:7px 0}.hero p{color:#C1CBE0;margin:0;max-width:720px}
.pill{display:inline-block;padding:5px 10px;border-radius:20px;background:#5b3df522;border:1px solid #806cff66;color:#D4CDFF;font-size:11px;font-weight:700}
.metric{background:linear-gradient(145deg,#192541,#131c32);border:1px solid var(--line);border-radius:16px;padding:16px;min-height:108px}
.metric-label{color:var(--muted);font-size:12px;font-weight:600}.metric-value{font:700 27px 'Space Grotesk';margin:7px 0 2px;color:var(--text)}.metric-note{color:#AAB7D3;font-size:11px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:19px}
.panel-title{font:600 16px 'Space Grotesk';margin-bottom:4px}.panel-sub{color:var(--muted);font-size:12px;margin-bottom:14px}
.stButton>button,.stDownloadButton>button{border:0;border-radius:10px;color:white;font-weight:700;background:linear-gradient(100deg,var(--purple),#397DFF);padding:.58rem 1rem}
.stButton>button:hover,.stDownloadButton>button:hover{border:0;filter:brightness(1.12);color:white}
.stTextInput input,.stTextArea textarea,.stSelectbox div[data-baseweb="select"]>div{background:#101A30!important;color:var(--text)!important;border-color:var(--line)!important;border-radius:10px}
[data-testid="stFileUploader"]{background:#101A30;border:1px dashed #7182AB;border-radius:14px;padding:12px}
[data-testid="stFileUploader"] *{color:var(--text)!important}
[data-testid="stAlert"]{border-radius:12px!important;background:#151D35!important;border:1px solid rgba(154,170,210,.24)!important;color:#F4F6FF!important}
[data-testid="stAlert"] *{color:#F4F6FF!important}
[data-testid="stAlert"] a{color:#8DBBFF!important}
[data-testid="stFileUploader"]{background:#101A30!important;border:1px dashed #7182AB!important;border-radius:14px!important;padding:12px!important}
[data-testid="stFileUploader"] section{background:#101A30!important}
[data-testid="stFileUploader"] button{background:#243354!important;color:#F4F6FF!important;border:1px solid #53688F!important}
[data-testid="stExpander"] details{background:#151D35!important;border:1px solid var(--line)!important;border-radius:12px!important}
[data-testid="stExpander"] summary,[data-testid="stExpander"] summary *{color:#F4F6FF!important}
.stTextArea textarea{background:#101A30!important;color:#F4F6FF!important;caret-color:#F4F6FF!important}
hr{border-color:var(--line)}footer{visibility:hidden}
</style>
""", unsafe_allow_html=True)

if "documents" not in st.session_state:
    st.session_state.documents = {}

def extract_text(name: str, raw: bytes):
    suffix = name.lower().rsplit(".", 1)[-1]
    if suffix == "txt":
        try:
            return raw.decode("utf-8-sig"), None
        except UnicodeDecodeError:
            return raw.decode("latin-1"), None
    if suffix == "pdf":
        reader = PdfReader(BytesIO(raw))
        pages = []
        for index, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            if text.strip():
                pages.append(f"[Page {index}]\n{text.strip()}")
        return "\n\n".join(pages), None
    if suffix == "docx":
        doc = Document(BytesIO(raw))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        tables = []
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    tables.append(" | ".join(cells))
        combined = "\n".join(paragraphs)
        if tables:
            combined += "\n\n[Tables]\n" + "\n".join(tables)
        return combined, None
    raise ValueError("Unsupported file type.")
