# GraphPilot AI — Complete Streamlit Build

This package replaces the current root-level app files with a consolidated build.

## Included
- Existing dark GraphPilot workspace and sidebar navigation
- Multi-file upload: PDF, DOCX, TXT, CSV, XLSX
- Document text extraction and searchable chunk retrieval
- Grounded document Q&A with source file/section references
- Gemini primary provider and Groq fallback (or selectable provider order)
- Groq model auto-detection: uses an accessible active Llama model when the API key can see one; otherwise uses `openai/gpt-oss-120b`
- CSV / Excel analytics: data quality, distributions, correlations, feature ranking
- Lightweight numeric regression preview with test-set metrics
- AI dataset insights
- AI report generation and TXT / PDF / PPTX export
- Settings page with live provider connection tests
- AI response card uses black text on white background for readability

## Existing Streamlit secrets
Your current secret names are retained. No need to paste the keys again if they are already present in Streamlit Cloud:

```toml
GEMINI_API_KEY = "your-existing-gemini-key"
GROQ_API_KEY = "your-existing-groq-key"
```

Optional model overrides:

```toml
GEMINI_MODEL = "gemini-3.8-flash"
GROQ_MODEL = "openai/gpt-oss-120b"
```

If `GROQ_MODEL` is not set, the app checks the models available to your Groq key and chooses an active Llama model if one is available to that account. If not, it uses the current Groq GPT-OSS 120B model as a fallback. Some older Groq Llama model IDs have been retired or restricted, so the app does not hard-code a retired model.

## GitHub replacement instructions
1. Download and extract this ZIP on your computer.
2. Open your existing repository `abrarhassanai573/graphpilot-ai` on GitHub.
3. Replace the existing `streamlit_app.py` and `requirements.txt` in the repository root with the files from this ZIP.
4. Upload the `.streamlit` folder too, including `.streamlit/config.toml`.
5. Also replace `.gitignore`, `.env.example`, and `README.md`.
6. Commit the changes to the `main` branch.
7. Streamlit Community Cloud should redeploy automatically. Keep the main file path as `streamlit_app.py`.

Do not upload the ZIP file itself as a repository source file. Upload the extracted files and folders.

## Important limitations
- Streamlit session memory is temporary. Uploaded documents, chat history, and datasets may be lost when the session expires or app restarts. Download reports you need to keep.
- Scanned/image-only PDFs need OCR, which is not included.
- The numeric prediction is an exploratory linear-regression baseline, not a production forecasting model.
- Permanent database storage, authentication/multi-user permissions, and Google Drive OAuth import are not included in this consolidated Streamlit package because they require separate credentials and infrastructure.
- AI usage depends on the quotas, model access, and billing/rate limits of the existing Gemini and Groq API accounts. The app automatically tries the second provider if the first one fails.

## Run locally (optional)
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Troubleshooting
- `401` / `403`: check whether the API key is valid and has access to the provider/model.
- `404 model not found`: model ID is unavailable to your account; remove `GROQ_MODEL` override so the app can auto-detect, or use a current model available to your account.
- `429`: provider rate limit/quota reached; automatic fallback tries the other provider.
- `ModuleNotFoundError`: confirm all `requirements.txt` dependencies were committed and wait for Streamlit to finish installing.
- `No module named ...`: check the Streamlit build logs; do not add packages to GitHub one at a time unless the error identifies a missing dependency.
- PDF has no extracted text: likely a scanned/image-only PDF; OCR is not included in this version.
