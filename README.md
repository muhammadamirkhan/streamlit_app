# Muraba Veil — Unit Manager

Interactive pricing & inventory tool for the Muraba Veil tower. Add / edit / remove
floors and units, adjust escalation and terrace parameters, and see portfolio value,
summary-by-type and topology statistics update live.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

The app is password protected. Locally the password is read from
`.streamlit/secrets.toml` (default `muraba2026`).

## Deploy (Streamlit Community Cloud)

1. Push this repo to GitHub.
2. On https://share.streamlit.io → **New app** → pick this repo, branch, and
   `streamlit_app.py` as the main file.
3. In **App → Settings → Secrets**, add:
   ```toml
   password = "your-chosen-password"
   ```
4. Deploy.

## Files

| File | Purpose |
|------|---------|
| `streamlit_app.py` | The application |
| `Muraba Veil Unit list.xlsx` | Source data |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme |
