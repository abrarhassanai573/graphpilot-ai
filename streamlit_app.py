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

# Keep styling scoped to Streamlit components. Avoid global html/body rules
# that can make the entire app appear blank in some browser environments.
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');

:root {
    --gp-bg: #0B1020;
    --gp-panel: #151D35;
    --gp-panel2: #192440;
    --gp-line: rgba(154,170,210,.20);
    --gp-muted: #AAB7D3;
    --gp-text: #F4F6FF;
    --gp-purple: #795CFF;
    --gp-blue: #168BFF;
}
[data-testid="stAppViewContainer"] {
    background: #0B1020;
}
[data-testid="stHeader"] {
    background: rgba(11,16,32,.92);
}
[data-testid="stSidebar"] {
    background: #0E162B;
    border-right: 1px solid var(--gp-line);
}
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}
h1, h2, h3 {
    font-family: 'Space Grotesk', sans-serif !important;
    color: #F4F6FF !important;
    letter-spacing: -.025em;
}
[data-testid="stMarkdownContainer"] p,
[data-testid="stWidgetLabel"] p,
[data-testid="stCaptionContainer"] {
    color: #E8ECFA;
}
.brand {
    display: flex;
    align-items: center;
    gap: 11px;
    padding: 8px 4px 22px;
}
.brand-mark {
    width: 40px;
    height: 40px;
    border-radius: 13px;
    display: grid;
    place-items: center;
    background: linear-gradient(135deg,#795CFF,#168BFF);
    font-size: 20px;
}
.brand-name {
    color: #F4F6FF;
    font: 700 19px 'Space Grotesk', sans-serif;
}
.brand-sub {
    color: #AAB7D3;
    font-size: 10px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
}
.eyebrow {
    color: #B6A9FF;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.7px;
    text-transform: uppercase;
}
.hero {
    padding: 25px 28px;
    border: 1px solid var(--gp-line);
    border-radius: 20px;
    background: linear-gradient(120deg,#172342,#111A31);
    margin-bottom: 20px;
}
.hero h1 {
    font-size: 32px;
    margin: 8px 0;
}
.hero p {
    color: #C1CBE0;
    margin: 0;
    max-width: 760px;
}
.metric {
    background: linear-gradient(145deg,#192541,#131C32);
    border: 1px solid var(--gp-line);
    border-radius: 16px;
    padding: 16px;
    min-height: 108px;
}
.metric-label {
    color: #AAB7D3;
    font-size: 12px;
    font-weight: 600;
}
.metric-value {
    color: #F4F6FF;
    font: 700 27px 'Space Grotesk', sans-serif;
    margin: 7px 0 2px;
}
.metric-note {
    color: #AAB7D3;
    font-size: 11px;
}
.stButton > button, .stDownloadButton > button {
    border: 0;
    border-radius: 10px;
    color: white;
    font-weight: 700;
    background: linear-gradient(100deg,#795CFF,#397DFF);
}
.stButton > button:hover, .stDownloadButton > button:hover {
    color: white;
    filter: brightness(1.1);
}
.stTextInput input, .stTextArea textarea {
    background: #101A30 !important;
    color: #F4F6FF !important;
    border-color: rgba(154,170,210,.25) !important;
}
[data-testid="stFileUploader"] {
    background: #101A30;
    border: 1px dashed #7182AB;
    border-radius: 14px;
    padding: 12px;
}
[data-testid="stFileUploader"] * {
    color: #F4F6FF !important;
}
[data-testid="stAlert"] {
    border-radius: 12px;
}
[data-testid="stExpander"] {
    background: #151D35;
    border: 1px solid rgba(154,170,210,.20);
    border-radius: 12px;
}
footer {
    visibility: hidden;
}


[data-testid="stAlert"] {
    background: #151D35 !important;
    border: 1px solid rgba(154, 170, 210, 0.25) !important;
    border-radius: 12px !important;
    color: #F4F6FF !important;
}
[data-testid="stAlert"] * {
    color: #F4F6FF !important;
}
[data-testid="stAlert"] svg {
    fill: currentColor;
}
</style>
""", unsafe_allow_html=True)

if "documents" not in st.session_state:
    st.session_state.documents = {}
if "assistant_history" not in st.session_state:
    st.session_state.assistant_history = []

def extract_text(name: str, raw: bytes) -> str:
    suffix = name.lower().rsplit(".", 1)[-1]
    if suffix == "txt":
        try:
            return raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            return raw.decode("latin-1")
    if suffix == "pdf":
        reader = PdfReader(BytesIO(raw))
        pages = []
        for page_number, page in enumerate(reader.pages, start=1):
            page_text = page.extract_text() or ""
            if page_text.strip():
                pages.append(f"[Page {page_number}]\n{page_text.strip()}")
        return "\n\n".join(pages)
    if suffix == "docx":
        doc = Document(BytesIO(raw))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        table_rows = []
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    table_rows.append(" | ".join(cells))
        combined = "\n".join(paragraphs)
        if table_rows:
            combined += "\n\n[Tables]\n" + "\n".join(table_rows)
        return combined
    raise ValueError("Unsupported file type. Please upload PDF, DOCX, or TXT.")

def make_pdf(title: str, content: str) -> bytes:
    output = BytesIO()
    pdf = SimpleDocTemplate(
        output, pagesize=A4, rightMargin=48, leftMargin=48,
        topMargin=48, bottomMargin=48
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph(title.replace("&", "&amp;"), styles["Title"]),
        Spacer(1, 0.2 * inch),
    ]
    for line in content.splitlines():
        safe_line = (line.replace("&", "&amp;")
                         .replace("<", "&lt;")
                         .replace(">", "&gt;"))
        story.append(Paragraph(safe_line if safe_line.strip() else "&nbsp;", styles["BodyText"]))
        story.append(Spacer(1, 5))
    pdf.build(story)
    return output.getvalue()

def make_pptx(title: str, content: str) -> bytes:
    presentation = Presentation()
    cover = presentation.slides.add_slide(presentation.slide_layouts[0])
    cover.shapes.title.text = title
    cover.placeholders[1].text = "Generated by GraphPilot AI"
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    chunks = [lines[i:i + 7] for i in range(0, len(lines), 7)]
    if not chunks:
        chunks = [["No text extracted."]]
    for slide_number, chunk in enumerate(chunks, start=1):
        slide = presentation.slides.add_slide(presentation.slide_layouts[1])
        slide.shapes.title.text = f"Document Content - {slide_number}"
        frame = slide.placeholders[1].text_frame
        frame.clear()
        for line_number, line in enumerate(chunk):
            paragraph = frame.paragraphs[0] if line_number == 0 else frame.add_paragraph()
            paragraph.text = line[:500]
            paragraph.font.size = Pt(16)
    output = BytesIO()
    presentation.save(output)
    return output.getvalue()

def show_hero(kicker: str, heading: str, description: str) -> None:
    st.markdown(
        f'<div class="hero"><div class="eyebrow">{kicker}</div>'
        f'<h1>{heading}</h1><p>{description}</p></div>',
        unsafe_allow_html=True,
    )

def show_metric(label: str, value: str, note: str) -> None:
    st.markdown(
        f'<div class="metric"><div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-note">{note}</div></div>',
        unsafe_allow_html=True,
    )

# Navigation
with st.sidebar:
    st.markdown(
        '<div class="brand"><div class="brand-mark">🧭</div><div>'
        '<div class="brand-name">GraphPilot AI</div>'
        '<div class="brand-sub">Document Intelligence</div></div></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="eyebrow">WORKSPACE</div>', unsafe_allow_html=True)
    page = st.radio(
        "Navigation",
        ["Overview", "Knowledge Base", "AI Assistant", "Reports", "Analytics", "Settings"],
        label_visibility="collapsed",
    )
    st.divider()
    st.caption("PHASE 2 · DOCUMENT INGESTION")
    st.caption("Uploaded files are held in this active session.")

documents = st.session_state.documents
total_chars = sum(len(doc["text"]) for doc in documents.values())
total_pages = sum(doc.get("pages", 0) for doc in documents.values())

if page == "Overview":
    show_hero(
        "DOCUMENT INTELLIGENCE PLATFORM",
        "Turn documents into decisions.",
        "Upload your knowledge, organize important information, and prepare your workspace for AI-powered analysis.",
    )
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        show_metric("DOCUMENTS", str(len(documents)), "Files in this session")
    with col2:
        show_metric("EXTRACTED CHARACTERS", f"{total_chars:,}", "Readable text extracted")
    with col3:
        show_metric("PDF PAGES", f"{total_pages:,}", "Pages detected")
    with col4:
        show_metric("AI PROVIDERS", "2", "Gemini and Groq configuration slots")

    st.subheader("Workspace")
    left, right = st.columns([1.4, 1])
    with left:
        st.markdown("#### Knowledge workspace")
        if documents:
            for item in list(documents.values())[:6]:
                st.write(f"📄 **{item['name']}** · {len(item['text']):,} characters")
        else:
            st.info("Your knowledge base is empty. Open Knowledge Base to upload your first document.")
    with right:
        st.markdown("#### Platform status")
        st.success("Document extraction is ready.")
        st.warning("Live AI and RAG answering are not connected in this phase.")
        st.caption("Next phase: embeddings, retrieval, citations, and provider-backed answers.")

elif page == "Knowledge Base":
    show_hero(
        "KNOWLEDGE BASE",
        "Bring your documents together.",
        "Upload PDF, DOCX, and TXT files. Preview extracted text and export it in multiple formats.",
    )
    uploaded_files = st.file_uploader(
        "Choose documents",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
        help="Scanned PDFs may not contain selectable text for extraction.",
    )
    if uploaded_files:
        new_count = 0
        for uploaded_file in uploaded_files:
            raw = uploaded_file.getvalue()
            file_hash = hashlib.sha256(raw).hexdigest()
            if file_hash in documents:
                continue
            try:
                extracted = extract_text(uploaded_file.name, raw)
                page_count = (
                    len(PdfReader(BytesIO(raw)).pages)
                    if uploaded_file.name.lower().endswith(".pdf")
                    else 0
                )
                documents[file_hash] = {
                    "name": uploaded_file.name,
                    "text": extracted,
                    "size": len(raw),
                    "pages": page_count,
                    "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
                new_count += 1
            except Exception as error:
                st.error(f"Could not process {uploaded_file.name}: {error}")
        if new_count:
            st.success(f"{new_count} document(s) added to your session.")

    st.subheader("Your documents")
    if not documents:
        st.info("No documents uploaded yet.")
    else:
        for file_hash, item in list(documents.items()):
            with st.expander(f"📄 {item['name']} · {len(item['text']):,} characters"):
                st.caption(
                    f"Uploaded: {item['uploaded_at']} · "
                    f"File size: {item['size']:,} bytes · PDF pages: {item['pages']}"
                )
                st.text_area(
                    "Extracted text preview",
                    item["text"][:12000],
                    height=220,
                    key=f"preview_{file_hash}",
                )
                filename = Path(item["name"]).stem
                txt_col, pdf_col, ppt_col, remove_col = st.columns(4)
                with txt_col:
                    st.download_button(
                        "Download TXT", item["text"].encode("utf-8"),
                        file_name=f"{filename}_extracted.txt",
                        mime="text/plain", key=f"txt_{file_hash}",
                    )
                with pdf_col:
                    st.download_button(
                        "Download PDF", make_pdf(filename, item["text"]),
                        file_name=f"{filename}_extracted.pdf",
                        mime="application/pdf", key=f"pdf_{file_hash}",
                    )
                with ppt_col:
                    st.download_button(
                        "Download PPTX", make_pptx(filename, item["text"]),
                        file_name=f"{filename}_extracted.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"pptx_{file_hash}",
                    )
                with remove_col:
                    if st.button("Remove", key=f"remove_{file_hash}"):
                        del documents[file_hash]
                        st.rerun()

elif page == "AI Assistant":
    show_hero(
        "AI ASSISTANT",
        "Ask your knowledge base.",
        "A multilingual assistant interface prepared for document-grounded conversations.",
    )
    response_language = st.selectbox(
        "Preferred response language", ["English", "Urdu", "Roman Urdu"]
    )
    st.info(
        "AI provider calls and Retrieval-Augmented Generation (RAG) are not enabled "
        "in Phase 2. This page does not yet generate answers from your documents."
    )
    question = st.text_area(
        "Your question",
        placeholder="Example: Summarize the key points in my uploaded document.",
    )
    if st.button("Prepare question", type="primary"):
        if not question.strip():
            st.warning("Please enter a question first.")
        elif not documents:
            st.warning("Upload at least one document in Knowledge Base first.")
        else:
            st.session_state.assistant_history.append(
                {
                    "question": question.strip(),
                    "language": response_language,
                    "time": datetime.now().strftime("%H:%M"),
                }
            )
            st.warning(
                "Your question was saved in this session. No AI answer was generated "
                "because the AI/RAG connection is not active yet."
            )
    if st.session_state.assistant_history:
        st.subheader("Recent questions")
        for entry in reversed(st.session_state.assistant_history[-5:]):
            st.markdown(f"**{entry['question']}**")
            st.caption(f"{entry['language']} · {entry['time']}")

elif page == "Reports":
    show_hero(
        "REPORT STUDIO",
        "Prepare clear document outputs.",
        "Export extracted document content. AI-generated reports will be added in a later phase.",
    )
    if not documents:
        st.info("Upload documents first to create exports.")
    else:
        selected_hash = st.selectbox(
            "Select a document",
            list(documents.keys()),
            format_func=lambda key: documents[key]["name"],
        )
        item = documents[selected_hash]
        st.write(f"**Selected:** {item['name']}")
        st.text_area(
            "Report content preview",
            item["text"][:12000],
            height=260,
            key=f"report_{selected_hash}",
        )
        filename = Path(item["name"]).stem
        txt_col, pdf_col, ppt_col = st.columns(3)
        with txt_col:
            st.download_button(
                "Export TXT", item["text"].encode("utf-8"),
                file_name=f"{filename}_report.txt", mime="text/plain",
            )
        with pdf_col:
            st.download_button(
                "Export PDF", make_pdf(f"{filename} - Extracted Report", item["text"]),
                file_name=f"{filename}_report.pdf", mime="application/pdf",
            )
        with ppt_col:
            st.download_button(
                "Export PowerPoint", make_pptx(filename, item["text"]),
                file_name=f"{filename}_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

elif page == "Analytics":
    show_hero(
        "WORKSPACE ANALYTICS",
        "Understand your document collection.",
        "A live overview of files processed in the current browser session.",
    )
    col1, col2, col3 = st.columns(3)
    with col1:
        show_metric("TOTAL FILES", str(len(documents)), "Documents loaded")
    with col2:
        show_metric("TOTAL TEXT", f"{total_chars:,}", "Extracted characters")
    with col3:
        show_metric("PDF PAGES", f"{total_pages:,}", "PDF pages detected")
    if documents:
        st.subheader("Document breakdown")
        for item in documents.values():
            st.write(f"**{item['name']}**")
            ratio = len(item["text"]) / max(total_chars, 1)
            st.progress(min(ratio, 1.0))
            st.caption(f"{len(item['text']):,} characters")

elif page == "Settings":
    show_hero(
        "SETTINGS",
        "Configure your workspace.",
        "Review provider secret availability and how this deployment handles uploaded files.",
    )
    st.subheader("AI provider configuration")
    try:
        gemini_ready = bool(st.secrets.get("GEMINI_API_KEY", ""))
    except Exception:
        gemini_ready = False
    try:
        groq_ready = bool(st.secrets.get("GROQ_API_KEY", ""))
    except Exception:
        groq_ready = False

    gemini_col, groq_col = st.columns(2)
    with gemini_col:
        st.markdown("**Google Gemini**")
        if gemini_ready:
            st.success("Secret detected")
        else:
            st.warning("GEMINI_API_KEY not configured")
    with groq_col:
        st.markdown("**Groq**")
        if groq_ready:
            st.success("Secret detected")
        else:
            st.warning("GROQ_API_KEY not configured")

    st.caption(
        "Add keys in Streamlit Community Cloud → App settings → Secrets. "
        "Never paste API keys into GitHub or this chat."
    )
    st.subheader("Data and privacy")
    st.info(
        "Uploaded documents are stored in Streamlit session memory only. "
        "They are not saved to a permanent database and may disappear when the session restarts."
    )
    st.subheader("Current phase")
    st.write(
        "Phase 2 includes document upload, text extraction, preview, and TXT/PDF/PPTX exports."
    )
    st.write(
        "Not yet included: persistent storage, authentication, OCR, embeddings, vector search, "
        "live Gemini/Groq calls, or source-cited RAG answers."
    )
