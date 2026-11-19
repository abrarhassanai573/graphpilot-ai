import hashlib
from datetime import datetime
from io import BytesIO

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
[data-testid="stAlert"]{border-radius:12px}
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


def secret_is_configured(name):
    """Check secret presence without displaying the secret value."""
    try:
        return bool(str(st.secrets.get(name, "")).strip())
    except Exception:
        return False


def make_pdf(title, content):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=0.65*inch,
                            leftMargin=0.65*inch, topMargin=0.65*inch,
                            bottomMargin=0.65*inch)
    styles = getSampleStyleSheet()
    story = [Paragraph(title.replace("&", "&amp;"), styles["Title"]), Spacer(1, 14)]
    for line in content.splitlines():
        safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        story.extend([Paragraph(safe if safe.strip() else " ", styles["BodyText"]), Spacer(1, 5)])
    doc.build(story)
    return buffer.getvalue()

def make_pptx(title, content):
    prs = Presentation()
    cover = prs.slides.add_slide(prs.slide_layouts[0])
    cover.shapes.title.text = title
    cover.placeholders[1].text = "Extracted text · GraphPilot AI"
    lines = content.splitlines()
    chunks = [lines[i:i+12] for i in range(0, len(lines), 12)] or [[""]]
    for index, chunk in enumerate(chunks, 1):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Extracted Content — {index}"
        frame = slide.placeholders[1].text_frame
        frame.clear()
        for j, line in enumerate(chunk):
            para = frame.paragraphs[0] if j == 0 else frame.add_paragraph()
            para.text = line[:450] if line else " "
            para.font.size = Pt(15)
    output = BytesIO()
    prs.save(output)
    return output.getvalue()

def format_bytes(size):
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.2f} MB"

with st.sidebar:
    st.markdown("""
    <div class="brand"><div class="brand-mark">✦</div>
    <div><div class="brand-name">GraphPilot AI</div><div class="brand-sub">Document intelligence</div></div></div>
    <div class="eyebrow">Workspace</div>
    """, unsafe_allow_html=True)
    page = st.radio("Navigation", ["Overview","Knowledge Base","AI Assistant","Reports","Analytics","Settings"], label_visibility="collapsed")
    st.markdown("---")
    st.markdown(f"""<div class="panel" style="padding:14px"><b style="font-size:13px">Workspace status</b>
    <div style="color:#AAB7D3;font-size:11px;margin-top:5px">Phase 2 · Document ingestion</div>
    <div style="margin-top:12px"><span class="pill">● UI operational</span></div></div>""", unsafe_allow_html=True)
    st.markdown("<br><div style='color:#8292B4;font-size:10px'>GRAPH PILOT AI · 2026</div>", unsafe_allow_html=True)

st.markdown(f"""<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:15px">
<div><div class="eyebrow">Enterprise workspace</div><div style="font-size:13px;color:#AAB7D3;margin-top:4px">{page} / GraphPilot AI</div></div>
<div class="pill">PHASE 2 · DOCUMENT INGESTION</div></div>""", unsafe_allow_html=True)

docs = st.session_state.documents
total_chars = sum(len(d["text"]) for d in docs.values())

if page == "Overview":
    st.markdown("""<div class="hero"><div class="eyebrow">Your knowledge, connected</div>
    <h1>Turn documents into<br>decisions.</h1>
    <p>Bring your business documents into one workspace. Extracted text is available for review and will become the foundation for semantic search and grounded AI answers in upcoming phases.</p>
    <div style="margin-top:17px"><span class="pill">Document ingestion · Phase 2</span></div></div>""", unsafe_allow_html=True)
    c = st.columns(4)
    metrics = [("Documents uploaded",str(len(docs)),"Files in this session"),("Text extracted",str(sum(1 for d in docs.values() if d["text"].strip())),"Successfully processed"),("Characters extracted",f"{total_chars:,}","Across current documents"),("AI reports","0","Report studio comes later")]
    for col,(label,value,note) in zip(c,metrics):
        with col: st.markdown(f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div><div class="metric-note">{note}</div></div>',unsafe_allow_html=True)
    st.write("")
    a,b=st.columns([1.25,1])
    with a:
        st.markdown('<div class="panel"><div class="panel-title">Recent knowledge sources</div><div class="panel-sub">Latest files added to this session</div>',unsafe_allow_html=True)
        if docs:
            for d in list(docs.values())[-5:][::-1]:
                st.markdown(f"📄 **{d['name']}** · {d['size']} · {d['pages']} page(s)")
        else: st.markdown("<span style='color:#AAB7D3'>No documents uploaded yet. Open Knowledge Base to add your first files.</span>",unsafe_allow_html=True)
        st.markdown("</div>",unsafe_allow_html=True)
    with b:
        st.markdown('<div class="panel"><div class="panel-title">Platform roadmap</div><div class="panel-sub">Building the intelligence layer step by step</div>',unsafe_allow_html=True)
        for n,title,status in [("01","Workspace foundation","Complete"),("02","Document upload & extraction","Current phase"),("03","Semantic chunking & embeddings","Upcoming"),("04","RAG assistant & citations","Upcoming"),("05","Reports & integrations","Upcoming")]:
            st.markdown(f"<div style='display:flex;gap:12px;margin:12px 0'><div class='pill'>{n}</div><div><b style='font-size:12px'>{title}</b><div style='font-size:10px;color:#AAB7D3'>{status}</div></div></div>",unsafe_allow_html=True)
        st.markdown("</div>",unsafe_allow_html=True)

elif page == "Knowledge Base":
    st.title("Knowledge Base")
    st.caption("Upload, extract, review, and manage your document sources.")
    st.markdown('<div class="panel"><div class="panel-title">Add documents</div><div class="panel-sub">Supported formats: PDF, DOCX, TXT · Multiple files supported</div>',unsafe_allow_html=True)
    uploads = st.file_uploader("Choose files to add", type=["pdf","docx","txt"], accept_multiple_files=True, key="phase2_uploader")
    if uploads:
        if st.button("Process selected documents"):
            added, skipped = 0, 0
            for file in uploads:
                raw = file.getvalue()
                digest = hashlib.sha256(raw).hexdigest()
                if digest in docs:
                    skipped += 1
                    continue
                try:
                    extracted, _ = extract_text(file.name, raw)
                    pages = len(PdfReader(BytesIO(raw)).pages) if file.name.lower().endswith(".pdf") else (1 if extracted.strip() else 0)
                    docs[digest] = {"name":file.name,"size":format_bytes(len(raw)),"bytes":len(raw),"text":extracted,"pages":pages,"uploaded":datetime.now().strftime("%Y-%m-%d %H:%M")}
                    added += 1
                except Exception as exc:
                    st.error(f"Could not process {file.name}: {exc}")
            if added: st.success(f"Successfully processed {added} document(s).")
            if skipped: st.info(f"Skipped {skipped} duplicate file(s).")
    st.markdown("</div>",unsafe_allow_html=True)
    st.write("")
    st.markdown("### Your documents")
    if not docs:
        st.info("No documents in this session yet. Upload PDF, DOCX, or TXT files above.")
    else:
        st.caption(f"{len(docs)} document(s) · {total_chars:,} extracted characters")
        for digest,d in list(docs.items()):
            with st.expander(f"📄 {d['name']}  ·  {d['size']}  ·  {d['pages']} page(s)"):
                st.caption(f"Added: {d['uploaded']} · SHA-256: {digest[:16]}…")
                if not d["text"].strip():
                    st.warning("No selectable text was extracted. This can happen with scanned/image-only PDFs.")
                else:
                    st.text_area("Extracted text preview",d["text"],height=220,key=f"preview_{digest}")
                    stem = Path(d["name"]).stem
                    txt_data = d["text"].encode("utf-8")
                    pdf_data = make_pdf(f"Extracted Text — {d['name']}", d["text"])
                    pptx_data = make_pptx(f"Extracted Text — {d['name']}", d["text"])
                    st.markdown("**Download extracted text**")
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.download_button("Download TXT", txt_data, file_name=f"{stem}_extracted.txt", mime="text/plain", key=f"txt_{digest}", use_container_width=True)
                    with c2:
                        st.download_button("Download PDF", pdf_data, file_name=f"{stem}_extracted.pdf", mime="application/pdf", key=f"pdf_{digest}", use_container_width=True)
                    with c3:
                        st.download_button("Download PowerPoint", pptx_data, file_name=f"{stem}_extracted.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"pptx_{digest}", use_container_width=True)
                if st.button("Remove document",key=f"remove_{digest}"):
                    del st.session_state.documents[digest]
                    st.rerun()
    st.warning("Phase 2 stores uploaded files and extracted text in the current Streamlit session only. Data may be lost when the app session restarts or expires. Do not upload confidential documents yet.")

elif page == "AI Assistant":
    st.title("AI Assistant")
    st.caption("Your document-aware assistant will be powered by retrieval in a later phase.")
    st.info("Documents can now be uploaded and text-extracted. Semantic search, Gemini/Groq responses, and source citations are not connected yet.")
    q=st.text_area("Your question",placeholder="What are the main points across my documents?")
    if st.button("Ask GraphPilot"):
        st.warning("The AI engine is not connected in Phase 2. Your question was not sent to an AI provider.")

elif page == "Reports":
    st.title("Reports")
    st.caption("Create decision-ready reports from your knowledge sources.")
    st.info("AI report generation, report preview and revision, PDF export, and PowerPoint export are planned for later phases.")
    st.markdown('<div class="panel"><div class="panel-title">Report studio</div><div class="panel-sub">Your generated reports will appear here.</div><h2>📑</h2><b>No reports generated yet</b></div>',unsafe_allow_html=True)

elif page == "Analytics":
    st.title("Analytics")
    st.caption("Understand how your organization uses its knowledge.")
    x,y,z=st.columns(3)
    x.metric("Documents in session",len(docs)); y.metric("Extracted characters",f"{total_chars:,}"); z.metric("AI queries","0")
    st.info("Analytics currently reflect only the active session. Persistent activity analytics will be added with database integration.")

elif page == "Settings":
    st.title("Workspace settings")
    st.caption("Manage your GraphPilot AI workspace.")
    st.markdown('<div class="panel"><div class="panel-title">Workspace profile</div><div class="panel-sub">Authentication and organization-level access controls are planned for a later phase.</div></div>',unsafe_allow_html=True)
    st.text_input("Workspace name",value="My Organization",disabled=True)
    st.text_input("Current version",value="Phase 2 — document ingestion",disabled=True)
    st.markdown("### AI provider secrets")
    st.caption("Add the following exact names in Streamlit Community Cloud → App → Settings → Secrets. Do not paste API keys into GitHub or chat.")
    for secret_name, provider in [("GEMINI_API_KEY", "Google Gemini"), ("GROQ_API_KEY", "Groq")]:
        configured = secret_is_configured(secret_name)
        status = "Configured" if configured else "Not configured"
        symbol = "🟢" if configured else "⚪"
        st.markdown(f"**{provider}** · `{secret_name}` · {symbol} {status}")
    st.info("Phase 2 only checks whether these secrets exist. Gemini and Groq API calls are not active until the AI integration phase.")
    st.warning("This phase does not provide authentication, tenant isolation, or persistent secure storage. Do not upload confidential business documents.")

st.markdown("<hr><div style='display:flex;justify-content:space-between;color:#8292B4;font-size:10px'><span>GRAPHPILOT AI · DOCUMENT INTELLIGENCE</span><span>PHASE 2 · DOCUMENT INGESTION</span></div>",unsafe_allow_html=True)
