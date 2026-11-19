# GraphPilot AI — Phase 1

A premium dark-themed foundation for an enterprise document intelligence workspace.

## Phase 1 scope

This phase provides the Streamlit application shell and polished workspace UI:

- Overview dashboard and phased roadmap
- Knowledge Base page with a non-persistent file-selection preview
- AI Assistant page placeholder
- Reports page placeholder
- Analytics page placeholder
- Workspace Settings page placeholder
- Responsive dark visual theme

## Important limitations

Phase 1 is a UI and deployment foundation only. It does **not** yet implement document parsing, persistent uploads, embeddings, vector search, RAG, AI provider calls, authentication, organization isolation, Google Drive integration, or report generation. Do not upload confidential business documents to this preview.

## Technology

- Python
- Streamlit
- Custom CSS

## Run locally

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

Install dependencies and start the app:

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy on Streamlit Community Cloud

1. Push the contents of this project folder to a GitHub repository.
2. Open https://share.streamlit.io/ and sign in with GitHub.
3. Select **Create app**.
4. Choose the repository and branch.
5. Set the main file path to `streamlit_app.py`.
6. Deploy.

Streamlit Community Cloud installs dependencies from the root `requirements.txt`.

## Project phases

Upcoming phases will add secure document ingestion and storage, topic-aware chunking, embeddings and retrieval, Gemini/Groq integrations, grounded Q&A with citations, reports, and integrations. Security and tenant boundaries must be implemented and tested before using real confidential documents.
