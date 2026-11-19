# GraphPilot AI — Phase 2

A premium dark-themed document intelligence workspace with multi-file upload, text extraction, preview, and export.

## Features
- Upload multiple PDF, DOCX, and TXT files
- Extract selectable PDF text, Word paragraphs/tables, and text files
- Preview extracted text in the application
- Download extracted text as TXT, PDF, or PowerPoint (PPTX)
- View file size, page count, upload time, and content hash
- Detect duplicate uploads within the active session
- Remove documents from the active session

## Important limitations
Uploaded files and extracted text are held in Streamlit session memory only. They are not persisted to a database or durable object storage and may be lost when the session expires or app restarts. Scanned/image-only PDFs may not yield text because OCR is not included. Authentication, organization isolation, semantic chunking, embeddings, RAG, AI provider calls, and AI report generation are not implemented yet. Do not upload confidential documents.

## Run locally
```bash
python -m venv .venv
```
Windows:
```bash
.venv\\Scripts\\activate
```
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy on Streamlit Community Cloud
1. Push the project files to the root of your GitHub repository.
2. In Streamlit Community Cloud, open the app settings.
3. Select the connected repository and branch.
4. Set the main file path to `streamlit_app.py`.
5. Save and redeploy.
