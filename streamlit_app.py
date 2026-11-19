import streamlit as st
from datetime import datetime
from pathlib import Path

st.set_page_config(
    page_title="GraphPilot AI | Document Intelligence",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Theme and styling ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
:root {
  --bg: #0B1020; --panel: #151D35; --panel2: #192440; --line: rgba(154,170,210,.14);
  --muted: #9AA8C7; --text: #F4F6FF; --purple: #795CFF; --blue: #168BFF;
}
html, body, [data-testid="stAppViewContainer"] {
  background: var(--bg); color: var(--text); font-family: 'DM Sans', sans-serif;
}
[data-testid="stHeader"] { background: rgba(11,16,32,.88); }
[data-testid="stSidebar"] {
  background: #0E162B; border-right: 1px solid var(--line);
}
[data-testid="stSidebar"] * { color: var(--text); }
.block-container { padding-top: 1.7rem; padding-bottom: 3rem; max-width: 1500px; }
h1,h2,h3 { font-family: 'Space Grotesk', sans-serif !important; letter-spacing: -.035em; color: var(--text); }
p, label, .stMarkdown { color: var(--text); }
small, .muted { color: var(--muted) !important; }
.brand { display:flex; align-items:center; gap:11px; padding: 8px 4px 24px; }
.brand-mark { width:38px; height:38px; border-radius:13px; display:grid; place-items:center;
  background:linear-gradient(135deg,var(--purple),var(--blue)); font-size:20px; box-shadow:0 8px 28px #5b3df544; }
.brand-name { font:700 19px 'Space Grotesk'; letter-spacing:-.7px; }
.brand-sub { color:var(--muted); font-size:10px; letter-spacing:1.7px; text-transform:uppercase; }
.eyebrow { color:#9C8BFF; font-size:11px; font-weight:700; letter-spacing:1.8px; text-transform:uppercase; }
.hero { padding:26px 28px; border:1px solid var(--line); border-radius:20px;
  background:radial-gradient(ellipse at 88% 0%, #5b3df533 0%, transparent 42%),
             linear-gradient(120deg,#172342,#111a31); margin-bottom:22px; }
.hero h1 { font-size:34px; margin:7px 0 8px; }
.hero p { color:#B7C3DF; margin:0; max-width:690px; }
.pill { display:inline-block; padding:5px 10px; border-radius:20px; background:#5b3df522;
  border:1px solid #806cff55; color:#C7BDFF; font-size:11px; font-weight:700; }
.metric { background:linear-gradient(145deg,#192541,#131c32); border:1px solid var(--line);
  border-radius:16px; padding:17px 18px; min-height:120px; }
.metric-label { color:var(--muted); font-size:12px; font-weight:600; }
.metric-value { color:var(--text); font:700 28px 'Space Grotesk'; margin:7px 0 3px; }
.metric-note { color:#8FA0C5; font-size:11px; }
.panel { background:var(--panel); border:1px solid var(--line); border-radius:16px; padding:20px; }
.panel-title { font:600 16px 'Space Grotesk'; margin-bottom:4px; }
.panel-sub { color:var(--muted); font-size:12px; margin-bottom:15px; }
.stButton > button, .stDownloadButton > button {
  border:0; border-radius:10px; color:white; font-weight:700;
  background:linear-gradient(100deg,var(--purple),#397DFF); padding:.62rem 1rem;
}
.stButton > button:hover, .stDownloadButton > button:hover { border:0; filter:brightness(1.12); color:white; }
.stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] > div {
  background:#101A30 !important; color:var(--text) !important; border-color:var(--line) !important; border-radius:10px;
}
[data-testid="stFileUploader"] { background:#101A30; border:1px dashed #52638C; border-radius:14px; padding:12px; }
[data-testid="stFileUploader"] * { color:var(--text) !important; }
div[data-testid="stMetric"] { background:var(--panel); border:1px solid var(--line); padding:15px; border-radius:14px; }
hr { border-color:var(--line); }
footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# ---------- Session state ----------
if "active_page" not in st.session_state:
    st.session_state.active_page = "Overview"
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []

# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("""
    <div class="brand">
      <div class="brand-mark">✦</div>
      <div><div class="brand-name">GraphPilot AI</div><div class="brand-sub">Document intelligence</div></div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('<div class="eyebrow">Workspace</div>', unsafe_allow_html=True)
    page = st.radio(
        "Navigation",
        ["Overview", "Knowledge Base", "AI Assistant", "Reports", "Analytics", "Settings"],
        label_visibility="collapsed",
        key="navigation",
    )
    st.session_state.active_page = page
    st.markdown("---")
    st.markdown('<div class="panel" style="padding:14px;">'
                '<div style="font-weight:700;font-size:13px;">Workspace status</div>'
                '<div style="color:#9AA8C7;font-size:11px;margin-top:5px;">Phase 1 · Foundation preview</div>'
                '<div style="margin-top:12px;"><span class="pill">● System ready</span></div>'
                '</div>', unsafe_allow_html=True)
    st.markdown("<br><div style='color:#7181A5;font-size:10px;'>GRAPH PILOT AI · 2026</div>", unsafe_allow_html=True)

# ---------- Header ----------
st.markdown(f"""
<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
  <div><div class="eyebrow">Enterprise workspace</div>
  <div style="font:600 13px 'DM Sans';color:#9AA8C7;margin-top:4px;">{page} / GraphPilot AI</div></div>
  <div class="pill">PHASE 1 · FOUNDATION</div>
</div>
""", unsafe_allow_html=True)

if page == "Overview":
    st.markdown("""
    <div class="hero">
      <div class="eyebrow">Your knowledge, connected</div>
      <h1>Turn documents into<br>decisions.</h1>
      <p>A unified workspace for understanding business documents, finding trusted answers, and transforming information into actionable intelligence.</p>
      <div style="margin-top:18px;"><span class="pill">AI-powered document intelligence</span></div>
    </div>
    """, unsafe_allow_html=True)

    cols = st.columns(4)
    metrics = [
        ("Documents indexed", "0", "Upload documents to get started"),
        ("Knowledge chunks", "0", "Semantic indexing coming soon"),
        ("Questions answered", "0", "Your AI activity will appear here"),
        ("Reports generated", "0", "Turn insights into deliverables"),
    ]
    for col, (label, value, note) in zip(cols, metrics):
        with col:
            st.markdown(f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div class="metric-note">{note}</div></div>', unsafe_allow_html=True)

    st.write("")
    left, right = st.columns([1.35, 1])
    with left:
        st.markdown('<div class="panel"><div class="panel-title">Workspace activity</div><div class="panel-sub">Your document intelligence activity will appear here.</div>', unsafe_allow_html=True)
        st.markdown("""
        <div style="display:flex;gap:13px;align-items:center;padding:13px 0;border-bottom:1px solid #ffffff12;">
          <div style="background:#5b3df522;border-radius:10px;padding:10px;">📄</div>
          <div><b style="font-size:13px;">Document ingestion</b><div style="color:#9AA8C7;font-size:11px;">Upload and organize your knowledge sources</div></div>
          <div style="margin-left:auto;color:#9AA8C7;font-size:11px;">Planned</div>
        </div>
        <div style="display:flex;gap:13px;align-items:center;padding:13px 0;border-bottom:1px solid #ffffff12;">
          <div style="background:#168bff22;border-radius:10px;padding:10px;">✦</div>
          <div><b style="font-size:13px;">Grounded AI answers</b><div style="color:#9AA8C7;font-size:11px;">Ask questions with source references</div></div>
          <div style="margin-left:auto;color:#9AA8C7;font-size:11px;">Planned</div>
        </div>
        <div style="display:flex;gap:13px;align-items:center;padding:13px 0;">
          <div style="background:#27c99b22;border-radius:10px;padding:10px;">📊</div>
          <div><b style="font-size:13px;">Decision-ready reports</b><div style="color:#9AA8C7;font-size:11px;">Create and refine business reports</div></div>
          <div style="margin-left:auto;color:#9AA8C7;font-size:11px;">Planned</div>
        </div></div>
        """, unsafe_allow_html=True)
    with right:
        st.markdown('<div class="panel"><div class="panel-title">Build roadmap</div><div class="panel-sub">A phased path to the complete platform</div>', unsafe_allow_html=True)
        steps = [("01","Workspace foundation","Current phase",True),("02","Document experience","Next milestone",False),("03","Retrieval intelligence","Planned",False),("04","AI reports & integrations","Planned",False)]
        for n, title, status, done in steps:
            color = "#795CFF" if done else "#52617F"
            st.markdown(f'<div style="display:flex;gap:12px;margin:14px 0;"><div style="width:30px;height:30px;border-radius:9px;background:{color}22;color:#BDB1FF;display:grid;place-items:center;font-size:11px;font-weight:700;">{n}</div><div><div style="font-size:12px;font-weight:700;">{title}</div><div style="font-size:10px;color:#9AA8C7;">{status}</div></div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

elif page == "Knowledge Base":
    st.title("Knowledge Base")
    st.caption("A central place for your business documents and knowledge sources.")
    st.info("Phase 1 preview: file processing, persistent storage, semantic chunking, and indexing will be enabled in upcoming phases.")
    uploaded = st.file_uploader("Choose documents", type=["pdf", "docx", "txt"], accept_multiple_files=True)
    if uploaded:
        st.session_state.uploaded_files = [f.name for f in uploaded]
        st.success(f"{len(uploaded)} file(s) selected. This preview does not upload or persist files yet.")
    if st.session_state.uploaded_files:
        st.markdown("#### Selected files")
        for name in st.session_state.uploaded_files:
            st.write("📄", name)
    else:
        st.markdown('<div class="panel"><div style="font-size:28px;">📚</div><div class="panel-title">Your knowledge base is ready for its first source</div><div class="panel-sub">PDF, DOCX and TXT support is planned for the next implementation phase.</div></div>', unsafe_allow_html=True)

elif page == "AI Assistant":
    st.title("AI Assistant")
    st.caption("Ask questions across your organization’s knowledge.")
    st.info("AI responses, RAG retrieval, citations, and Gemini/Groq provider configuration will be implemented in upcoming phases.")
    prompt = st.text_area("Your question", placeholder="e.g. What are the key risks mentioned across the project documents?")
    if st.button("Ask GraphPilot", use_container_width=False):
        if prompt.strip():
            st.warning("The AI engine is not connected in Phase 1. Your question has not been sent to an AI provider.")
        else:
            st.error("Please enter a question.")

elif page == "Reports":
    st.title("Reports")
    st.caption("Create clear, decision-ready reports from your knowledge sources.")
    st.info("Report generation, preview, revision, PDF export, and PowerPoint export are planned for later phases.")
    st.markdown('<div class="panel"><div class="panel-title">Report studio</div><div class="panel-sub">Your generated reports will be organized here.</div><div style="font-size:34px;">📑</div><b>No reports yet</b><div style="color:#9AA8C7;font-size:12px;">Connect your documents and AI workflow to create your first report.</div></div>', unsafe_allow_html=True)

elif page == "Analytics":
    st.title("Analytics")
    st.caption("Understand how your organization uses its knowledge.")
    st.info("Live analytics will be connected as document processing and AI activity become available.")
    a, b, c = st.columns(3)
    a.metric("Knowledge sources", "0")
    b.metric("Queries this month", "0")
    c.metric("Active workspaces", "1")
    st.markdown('<div class="panel"><div class="panel-title">Usage overview</div><div class="panel-sub">No activity data available in the foundation phase.</div></div>', unsafe_allow_html=True)

elif page == "Settings":
    st.title("Workspace settings")
    st.caption("Manage your GraphPilot AI workspace configuration.")
    st.markdown('<div class="panel"><div class="panel-title">Workspace profile</div><div class="panel-sub">Foundation preview — authentication and organization controls are not active yet.</div></div>', unsafe_allow_html=True)
    st.text_input("Workspace name", value="My Organization", disabled=True)
    st.text_input("Plan", value="Development preview", disabled=True)
    st.warning("Do not upload confidential business documents in this preview. Authentication, tenant isolation, and secure persistent storage are not implemented yet.")

st.markdown("<hr><div style='display:flex;justify-content:space-between;color:#7181A5;font-size:10px;'><span>GRAPHPILOT AI · DOCUMENT INTELLIGENCE</span><span>PHASE 1 FOUNDATION</span></div>", unsafe_allow_html=True)
