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
    response_language = st.selectbox(
        "Preferred response language",
        ["Match the language of my question", "English", "Urdu", "Roman Urdu"],
        help="The upcoming AI integration will use this preference for its replies."
    )
    q=st.text_area("Your question",placeholder="Ask in English, Urdu, or Roman Urdu...")
    st.caption("Voice input and spoken answers will be added with the AI assistant integration.")
    if st.button("Ask GraphPilot"):
        st.warning("The AI engine is not connected in this phase. Your question was not sent to an AI provider. The next phase will enable document-grounded answers in English, Urdu, and Roman Urdu.")

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
